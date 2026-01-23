"""
Streaming utilities for Open WebUI.

This module provides heartbeat-aware streaming to prevent proxy timeouts
during long-running LLM operations.

Usage:
    from open_webui.utils.streaming import create_streaming_response_with_heartbeats
    
    # Wrap any async generator with heartbeats
    return create_streaming_response_with_heartbeats(
        body_iterator=my_async_generator,
        initial_message="Processing request..."
    )
"""

import asyncio
import json
import logging
from typing import AsyncGenerator, Optional, Any, Union

from starlette.responses import StreamingResponse

from open_webui.env import SSE_HEARTBEAT_SECONDS

log = logging.getLogger(__name__)


async def stream_with_heartbeats(
    async_generator: AsyncGenerator,
    heartbeat_interval: int = None,
    initial_message: str = None
) -> AsyncGenerator[bytes, None]:
    """
    Wraps an async generator to inject SSE heartbeat comments at regular intervals.
    
    This prevents proxy timeouts (e.g., Cloudflare's 100-second limit, Azure's 230-second
    limit) during long-running operations like:
    - Tool calls (web search, RAG retrieval)
    - Slow model responses (reasoning models like o1, QwQ)
    - Large document processing
    
    SSE comment format `: keep-alive\\n\\n` is ignored by SSE clients per the spec,
    but keeps the HTTP connection alive for proxies.
    
    Args:
        async_generator: The original async generator yielding SSE data
        heartbeat_interval: Seconds between heartbeat comments (default from SSE_HEARTBEAT_SECONDS env)
        initial_message: Optional initial status message to send immediately
    
    Yields:
        SSE formatted data with interleaved heartbeat comments as bytes
    """
    if heartbeat_interval is None:
        heartbeat_interval = SSE_HEARTBEAT_SECONDS
    
    # Send initial message immediately to open the stream and get first byte out
    if initial_message:
        yield f"data: {json.dumps({'status': initial_message})}\n\n".encode("utf-8")
    else:
        # Send a comment to immediately open the stream
        yield b": stream-start\n\n"
    
    # Create a queue to receive items from the generator
    queue: asyncio.Queue = asyncio.Queue()
    generator_done = asyncio.Event()
    generator_error: Optional[Exception] = None
    
    async def producer():
        """Consume the async generator and put items in queue"""
        nonlocal generator_error
        try:
            async for item in async_generator:
                await queue.put(item)
        except Exception as e:
            generator_error = e
            log.error(f"Error in stream producer: {e}")
        finally:
            generator_done.set()
    
    # Start the producer task
    producer_task = asyncio.create_task(producer())
    
    try:
        while not generator_done.is_set() or not queue.empty():
            try:
                # Wait for item with timeout for heartbeat
                item = await asyncio.wait_for(
                    queue.get(),
                    timeout=heartbeat_interval
                )
                
                # Yield the actual item
                if isinstance(item, bytes):
                    yield item
                elif isinstance(item, str):
                    yield item.encode("utf-8")
                else:
                    # Assume it's already properly formatted
                    yield item
                    
            except asyncio.TimeoutError:
                # Send heartbeat comment (ignored by SSE clients per spec)
                yield b": keep-alive\n\n"
        
        # Check if there was an error in the producer
        if generator_error:
            error_msg = json.dumps({"error": {"detail": str(generator_error)}})
            yield f"data: {error_msg}\n\n".encode("utf-8")
                
    except asyncio.CancelledError:
        log.debug("Stream cancelled")
        producer_task.cancel()
        raise
    except Exception as e:
        log.error(f"Error in heartbeat stream: {e}")
        error_msg = json.dumps({"error": {"detail": str(e)}})
        yield f"data: {error_msg}\n\n".encode("utf-8")
    finally:
        if not producer_task.done():
            producer_task.cancel()
            try:
                await producer_task
            except asyncio.CancelledError:
                pass


def create_streaming_response_with_heartbeats(
    body_iterator: AsyncGenerator,
    headers: dict = None,
    status_code: int = 200,
    background: Any = None,
    initial_message: str = None,
    heartbeat_interval: int = None
) -> StreamingResponse:
    """
    Creates a StreamingResponse with heartbeat injection and proper headers
    to prevent proxy buffering.
    
    Args:
        body_iterator: The async generator to wrap
        headers: Optional headers dict (will add anti-buffering headers)
        status_code: HTTP status code
        background: Background task to run after response
        initial_message: Optional initial status message
        heartbeat_interval: Optional custom heartbeat interval
    
    Returns:
        StreamingResponse with heartbeats and anti-buffering headers
    """
    # Ensure headers dict exists
    if headers is None:
        headers = {}
    
    # Add headers to prevent proxy buffering
    # These work with nginx, Cloudflare, Azure App Gateway, etc.
    headers.update({
        "Content-Type": "text/event-stream; charset=utf-8",
        "Cache-Control": "no-cache, no-transform",
        "Connection": "keep-alive",
        "X-Accel-Buffering": "no",           # Disable nginx buffering
        "X-Content-Type-Options": "nosniff",
    })
    
    return StreamingResponse(
        stream_with_heartbeats(
            body_iterator,
            heartbeat_interval=heartbeat_interval,
            initial_message=initial_message
        ),
        status_code=status_code,
        headers=headers,
        background=background,
        media_type="text/event-stream"
    )


def add_anti_buffering_headers(headers: dict = None) -> dict:
    """
    Adds anti-buffering headers to an existing headers dict.
    Useful when you want to add headers without wrapping the stream.
    
    Args:
        headers: Existing headers dict (will be modified in place if provided)
    
    Returns:
        Headers dict with anti-buffering headers added
    """
    if headers is None:
        headers = {}
    
    headers.update({
        "X-Accel-Buffering": "no",           # Disable nginx buffering
        "Cache-Control": "no-cache, no-transform",
        "Connection": "keep-alive",
    })
    
    return headers


# Convenience alias
create_heartbeat_streaming_response = create_streaming_response_with_heartbeats