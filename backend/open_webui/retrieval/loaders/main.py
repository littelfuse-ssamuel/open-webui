import requests
import logging
import ftfy
import sys
import json
from collections.abc import Iterable

from azure.identity import DefaultAzureCredential
from langchain_community.document_loaders import (
    AzureAIDocumentIntelligenceLoader,
    BSHTMLLoader,
    CSVLoader,
    Docx2txtLoader,
    OutlookMessageLoader,
    PyPDFLoader,
    TextLoader,
    UnstructuredEPubLoader,
    UnstructuredExcelLoader,
    UnstructuredODTLoader,
    UnstructuredPowerPointLoader,
    UnstructuredRSTLoader,
    UnstructuredXMLLoader,
    YoutubeLoader,
)
from langchain_core.documents import Document

from open_webui.retrieval.loaders.external_document import ExternalDocumentLoader
from open_webui.retrieval.loaders.mistral import MistralLoader
from open_webui.retrieval.loaders.datalab_marker import DatalabMarkerLoader
from open_webui.retrieval.loaders.mineru import MinerULoader

from open_webui.env import GLOBAL_LOG_LEVEL, REQUESTS_VERIFY

logging.basicConfig(stream=sys.stdout, level=GLOBAL_LOG_LEVEL)
log = logging.getLogger(__name__)


# Using set for O(1) lookup performance
known_source_ext = {
    "go",
    "py",
    "java",
    "sh",
    "bat",
    "ps1",
    "cmd",
    "js",
    "ts",
    "css",
    "cpp",
    "hpp",
    "h",
    "c",
    "cs",
    "sql",
    "log",
    "ini",
    "pl",
    "pm",
    "r",
    "dart",
    "dockerfile",
    "env",
    "php",
    "hs",
    "hsc",
    "lua",
    "nginxconf",
    "conf",
    "m",
    "mm",
    "plsql",
    "perl",
    "rb",
    "rs",
    "db2",
    "scala",
    "bash",
    "swift",
    "vue",
    "svelte",
    "ex",
    "exs",
    "erl",
    "tsx",
    "jsx",
    "lhs",
    "json",
    "cfg",
    "cfm",
    "cgi",
    "config",
    "csv",
    "html",
    "kt",
    "kts",
    "lisp",
    "md",
    "toml",
    "tsv",
    "txt",
    "vb",
    "xml",
    "yaml",
    "yml",
}


class TikaLoader:
    def __init__(self, url, file_path, mime_type=None, extract_images=None):
        self.url = url
        self.file_path = file_path
        self.mime_type = mime_type
        self.extract_images = extract_images

    def load(self) -> list[Document]:
        with open(self.file_path, "rb") as f:
            data = f.read()

        if self.mime_type is not None:
            headers = {"Content-Type": self.mime_type}
        else:
            headers = {}

        if self.extract_images == True:
            headers["X-Tika-PDFextractInlineImages"] = "true"

        endpoint = self.url
        if not endpoint.endswith("/"):
            endpoint += "/"
        endpoint += "tika/text"

        r = requests.put(endpoint, data=data, headers=headers, verify=REQUESTS_VERIFY)

        if r.ok:
            raw_metadata = r.json()
            text = raw_metadata.get("X-TIKA:content", "<No text content found>").strip()

            if "Content-Type" in raw_metadata:
                headers["Content-Type"] = raw_metadata["Content-Type"]

            log.debug("Tika extracted text: %s", text)

            return [Document(page_content=text, metadata=headers)]
        else:
            raise Exception(f"Error calling Tika: {r.reason}")


class DoclingLoader:
    """
    DoclingLoader with extended parameter support.
    
    Supports both simple usage and advanced configuration including:
    - OCR settings (do_ocr, force_ocr, ocr_engine, ocr_lang)
    - Picture description (local and API modes)
    - PDF backend selection
    - Table extraction modes
    - Pipeline configuration
    
    Littelfuse extended implementation.
    """

    def __init__(self, url, api_key=None, file_path=None, mime_type=None, params=None):
        self.url = url.rstrip("/")
        self.api_key = api_key
        self.file_path = file_path
        self.mime_type = mime_type
        self.params = params or {}

    def load(self) -> list[Document]:
        with open(self.file_path, "rb") as f:
            headers = {}
            if self.api_key:
                headers["X-Api-Key"] = f"Bearer {self.api_key}"

            files = {
                "files": (
                    self.file_path,
                    f,
                    self.mime_type or "application/octet-stream",
                )
            }

            # Build request params
            request_params = {"image_export_mode": "placeholder"}

            if self.params:
                # Picture description settings
                if self.params.get("do_picture_description"):
                    request_params["do_picture_description"] = self.params.get(
                        "do_picture_description"
                    )

                    picture_description_mode = self.params.get(
                        "picture_description_mode", ""
                    ).lower()

                    if picture_description_mode == "local" and self.params.get(
                        "picture_description_local", {}
                    ):
                        request_params["picture_description_local"] = json.dumps(
                            self.params.get("picture_description_local", {})
                        )
                    elif picture_description_mode == "api" and self.params.get(
                        "picture_description_api", {}
                    ):
                        request_params["picture_description_api"] = json.dumps(
                            self.params.get("picture_description_api", {})
                        )

                # OCR settings
                if self.params.get("do_ocr") is not None:
                    request_params["do_ocr"] = self.params.get("do_ocr")

                if self.params.get("force_ocr") is not None:
                    request_params["force_ocr"] = self.params.get("force_ocr")

                if (
                    self.params.get("do_ocr")
                    and self.params.get("ocr_engine")
                    and self.params.get("ocr_lang")
                ):
                    request_params["ocr_engine"] = self.params.get("ocr_engine")
                    request_params["ocr_lang"] = [
                        lang.strip()
                        for lang in self.params.get("ocr_lang").split(",")
                        if lang.strip()
                    ]

                # Additional settings
                if self.params.get("pdf_backend"):
                    request_params["pdf_backend"] = self.params.get("pdf_backend")

                if self.params.get("table_mode"):
                    request_params["table_mode"] = self.params.get("table_mode")

                if self.params.get("pipeline"):
                    request_params["pipeline"] = self.params.get("pipeline")

            r = requests.post(
                f"{self.url}/v1/convert/file",
                files=files,
                data=request_params,
                headers=headers,
            )

        if r.ok:
            result = r.json()
            document_data = result.get("document", {})
            text = document_data.get("md_content", "<No text content found>")

            metadata = {"Content-Type": self.mime_type} if self.mime_type else {}

            log.debug("Docling extracted text: %s", text)
            return [Document(page_content=text, metadata=metadata)]
        else:
            error_msg = f"Error calling Docling API: {r.reason}"
            if r.text:
                try:
                    error_data = r.json()
                    if "detail" in error_data:
                        error_msg += f" - {error_data['detail']}"
                except Exception:
                    error_msg += f" - {r.text}"
            raise Exception(f"Error calling Docling: {error_msg}")


class Loader:
    """
    Main document loader that dispatches to appropriate engine-specific loaders.
    
    Supports returning image references when using loaders that extract images
    (e.g., ExternalDocumentLoader with extract_images=True).
    
    Littelfuse extended implementation with image_refs support.
    """

    def __init__(self, engine: str = "", **kwargs):
        self.engine = engine
        self.user = kwargs.get("user", None)
        self.kwargs = kwargs

    def load(
        self, filename: str, file_content_type: str, file_path: str
    ) -> list[Document] | tuple[list[Document], list[str]]:
        """
        Load a document and optionally extract images.
        
        Returns:
            Either a list of Documents, or a tuple of (documents, image_refs)
            if the loader supports image extraction.
        """
        loader = self._get_loader(filename, file_content_type, file_path)
        log.info(
            f"Loader.load: engine='{self.engine}', loader_type={type(loader).__name__}"
        )
        raw_result = loader.load()
        log.info(
            f"Loader.load: raw_result type={type(raw_result)}, is_tuple={isinstance(raw_result, tuple)}"
        )

        # Check if result includes image_refs (tuple format from ExternalDocumentLoader)
        has_image_refs = isinstance(raw_result, tuple) and len(raw_result) == 2
        if has_image_refs:
            docs, image_refs = raw_result
            log.info(
                f"Loader.load: Extracted tuple with {len(docs)} docs and {len(image_refs)} image_refs"
            )
        else:
            docs = raw_result
            image_refs = []
            log.info(f"Loader.load: Non-tuple result, raw_result type: {type(raw_result)}")

        # Flatten nested document structures
        flat_docs: list[Document] = []

        def _flatten(items):
            for item in items:
                # Avoid treating strings/bytes/Documents as generic iterables
                if isinstance(item, Iterable) and not isinstance(
                    item, (str, bytes, Document)
                ):
                    yield from _flatten(item)
                else:
                    yield item

        for item in _flatten(docs):
            if isinstance(item, Document):
                flat_docs.append(
                    Document(
                        page_content=ftfy.fix_text(item.page_content),
                        metadata=item.metadata,
                    )
                )
            else:
                log.warning(
                    "Loader returned non-Document item of type %s; skipping: %r",
                    type(item),
                    item,
                )

        # Return tuple if we originally had image_refs
        if has_image_refs:
            log.info(
                f"Loader.load: Returning tuple with {len(flat_docs)} docs and {len(image_refs)} image_refs"
            )
            return flat_docs, image_refs

        log.info(f"Loader.load: Returning list with {len(flat_docs)} docs (no image_refs)")
        return flat_docs

    def _is_text_file(self, file_ext: str, file_content_type: str) -> bool:
        return file_ext in known_source_ext or (
            file_content_type
            and file_content_type.find("text/") >= 0
            # Avoid text/html files being detected as text
            and not file_content_type.find("html") >= 0
        )

    def _get_loader(self, filename: str, file_content_type: str, file_path: str):
        file_ext = filename.split(".")[-1].lower()
        # Normalize engine to lowercase for consistent matching
        engine = (self.engine or "").strip().lower()
        log.info(
            f"_get_loader: original_engine='{self.engine}', normalized_engine='{engine}', "
            f"file_ext='{file_ext}', content_type='{file_content_type}'"
        )

        # ===========================================
        # Engine-specific loaders
        # ===========================================

        if engine == "youtube":
            loader = YoutubeLoader.from_youtube_url(
                file_path, add_video_info=True, language="en"
            )

        elif engine in ("web", "external"):
            # ExternalDocumentLoader with image extraction support
            log.info(
                f"_get_loader: Using ExternalDocumentLoader with extract_images={self.kwargs.get('PDF_EXTRACT_IMAGES')}"
            )
            loader = ExternalDocumentLoader(
                url=self.kwargs.get("EXTERNAL_DOCUMENT_LOADER_URL"),
                api_key=self.kwargs.get("EXTERNAL_DOCUMENT_LOADER_API_KEY"),
                file_path=file_path,
                mime_type=file_content_type,
                user=self.user,
                extract_images=self.kwargs.get("PDF_EXTRACT_IMAGES"),
            )

        elif engine == "azure_document_intelligence" or engine == "document_intelligence":
            # Support both naming conventions
            endpoint = self.kwargs.get(
                "AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"
            ) or self.kwargs.get("DOCUMENT_INTELLIGENCE_ENDPOINT")
            api_key = self.kwargs.get(
                "AZURE_DOCUMENT_INTELLIGENCE_API_KEY"
            ) or self.kwargs.get("DOCUMENT_INTELLIGENCE_KEY")
            model = self.kwargs.get(
                "AZURE_DOCUMENT_INTELLIGENCE_MODEL"
            ) or self.kwargs.get("DOCUMENT_INTELLIGENCE_MODEL")

            if api_key and api_key != "":
                loader = AzureAIDocumentIntelligenceLoader(
                    file_path=file_path,
                    api_endpoint=endpoint,
                    api_key=api_key,
                    api_model=model,
                )
            else:
                loader = AzureAIDocumentIntelligenceLoader(
                    file_path=file_path,
                    api_endpoint=endpoint,
                    azure_credential=DefaultAzureCredential(),
                    api_model=model,
                )

        elif engine == "tika":
            tika_url = self.kwargs.get("TIKA_URL") or self.kwargs.get("TIKA_SERVER_URL")
            if self._is_text_file(file_ext, file_content_type):
                loader = TextLoader(file_path, autodetect_encoding=True)
            else:
                loader = TikaLoader(
                    url=tika_url,
                    file_path=file_path,
                    mime_type=file_content_type,
                    extract_images=self.kwargs.get("PDF_EXTRACT_IMAGES"),
                )

        elif engine == "docling":
            docling_url = self.kwargs.get("DOCLING_URL") or self.kwargs.get(
                "DOCLING_SERVER_URL"
            )
            if self._is_text_file(file_ext, file_content_type):
                loader = TextLoader(file_path, autodetect_encoding=True)
            else:
                # Parse DOCLING_PARAMS if it's a string
                params = self.kwargs.get("DOCLING_PARAMS", {})
                if not isinstance(params, dict):
                    try:
                        params = json.loads(params)
                    except json.JSONDecodeError:
                        log.error("Invalid DOCLING_PARAMS format, expected JSON object")
                        params = {}

                loader = DoclingLoader(
                    url=docling_url,
                    api_key=self.kwargs.get("DOCLING_API_KEY"),
                    file_path=file_path,
                    mime_type=file_content_type,
                    params=params,
                )

        elif engine in ("mistral", "mistral_ocr"):
            # Support both naming conventions
            base_url = self.kwargs.get("MISTRAL_OCR_API_BASE_URL")
            api_key = self.kwargs.get("MISTRAL_API_KEY") or self.kwargs.get(
                "MISTRAL_OCR_API_KEY"
            )

            loader = MistralLoader(
                base_url=base_url,
                api_key=api_key,
                file_path=file_path,
            )

        elif engine == "datalab_marker":
            api_key = self.kwargs.get("DATALAB_MARKER_API_KEY")
            api_base_url = self.kwargs.get("DATALAB_MARKER_API_BASE_URL", "")
            if not api_base_url or api_base_url.strip() == "":
                api_base_url = "https://www.datalab.to/api/v1/marker"

            loader = DatalabMarkerLoader(
                file_path=file_path,
                api_key=api_key,
                api_base_url=api_base_url,
                additional_config=self.kwargs.get("DATALAB_MARKER_ADDITIONAL_CONFIG"),
                use_llm=self.kwargs.get("DATALAB_MARKER_USE_LLM", False),
                skip_cache=self.kwargs.get("DATALAB_MARKER_SKIP_CACHE", False),
                force_ocr=self.kwargs.get("DATALAB_MARKER_FORCE_OCR", False),
                paginate=self.kwargs.get("DATALAB_MARKER_PAGINATE", False),
                strip_existing_ocr=self.kwargs.get(
                    "DATALAB_MARKER_STRIP_EXISTING_OCR", False
                ),
                disable_image_extraction=self.kwargs.get(
                    "DATALAB_MARKER_DISABLE_IMAGE_EXTRACTION", False
                ),
                format_lines=self.kwargs.get("DATALAB_MARKER_FORMAT_LINES", False),
                output_format=self.kwargs.get(
                    "DATALAB_MARKER_OUTPUT_FORMAT", "markdown"
                ),
            )

        elif engine == "mineru":
            mineru_timeout = self.kwargs.get("MINERU_API_TIMEOUT", 300)
            if mineru_timeout:
                try:
                    mineru_timeout = int(mineru_timeout)
                except ValueError:
                    mineru_timeout = 300

            loader = MinerULoader(
                file_path=file_path,
                api_mode=self.kwargs.get("MINERU_API_MODE", "local"),
                api_url=self.kwargs.get("MINERU_API_URL")
                or self.kwargs.get("MINERU_LOCAL_API_URL", "http://localhost:8000"),
                api_key=self.kwargs.get("MINERU_API_KEY", ""),
                params=self.kwargs.get("MINERU_PARAMS", {}),
                timeout=mineru_timeout,
            )

        else:
            # ===========================================
            # No specific engine matched - file-type-specific loaders
            # WARNING: These loaders do NOT return image_refs tuple
            # ===========================================
            log.warning(
                f"_get_loader: No engine matched (engine='{engine}'), using file-type-specific "
                f"loader for ext='{file_ext}'. Image extraction will NOT return image_refs."
            )

            if file_ext == "pdf":
                loader = PyPDFLoader(
                    file_path, extract_images=self.kwargs.get("PDF_EXTRACT_IMAGES")
                )
            elif file_ext == "csv":
                loader = CSVLoader(file_path, autodetect_encoding=True)
            elif file_ext == "rst":
                loader = UnstructuredRSTLoader(file_path, mode="elements")
            elif file_ext == "xml":
                loader = UnstructuredXMLLoader(file_path)
            elif file_ext in ["htm", "html"]:
                loader = BSHTMLLoader(file_path, open_encoding="unicode_escape")
            elif file_ext == "md":
                loader = TextLoader(file_path, autodetect_encoding=True)
            elif file_content_type == "application/epub+zip":
                loader = UnstructuredEPubLoader(file_path)
            elif (
                file_content_type
                == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                or file_ext == "docx"
            ):
                loader = Docx2txtLoader(file_path)
            elif file_content_type in [
                "application/vnd.ms-excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ] or file_ext in ["xls", "xlsx"]:
                loader = UnstructuredExcelLoader(file_path)
            elif file_content_type in [
                "application/vnd.ms-powerpoint",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ] or file_ext in ["ppt", "pptx"]:
                loader = UnstructuredPowerPointLoader(file_path)
            elif file_ext == "msg":
                loader = OutlookMessageLoader(file_path)
            elif file_ext == "odt":
                loader = UnstructuredODTLoader(file_path)
            elif self._is_text_file(file_ext, file_content_type):
                loader = TextLoader(file_path, autodetect_encoding=True)
            else:
                loader = TextLoader(file_path, autodetect_encoding=True)

        return loader