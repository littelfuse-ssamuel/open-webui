# Open WebUI Agent Policy

## Project Scope

- Treat this repository as a company-customized fork of Open WebUI.
- Prioritize minimal diffs and upstream mergeability.
- Do not perform unrelated refactors.

## Change Classes and Approval Gates

- Auto-apply scoped edits for:
- Requested bug fixes.
- Small, local feature updates.
- Tests or docs directly tied to the requested change.
- Require explicit user confirmation before:
- Database schema changes or migrations.
- Auth, RBAC, permissions, or other security-sensitive logic.
- API contract changes consumed by external clients.
- Dependency major-version upgrades.
- Environment or config default changes that can affect deploy/runtime behavior.

## Safety Constraints

- Never introduce or expose secrets, keys, or credentials.
- Preserve existing RBAC/auth behavior unless explicitly requested to change it.
- Avoid broad behavior changes outside the requested scope.
- If unexpected unrelated file changes are detected, stop and ask the user how to proceed.

## Implementation Discipline

- Edit only files required for the requested outcome.
- Keep local naming, style, and structure consistent with nearby code.
- Prefer targeted checks mapped to touched subsystems.

## Validation Matrix

- Frontend touched: run `npm run check`.
- Frontend logic tests touched: run `npm run test:frontend -- <pattern>` when applicable.
- Backend Python touched: run `pytest -q backend/open_webui/test`.
- Backend-only edits: also run `python -m compileall backend/open_webui` for quick syntax safety.
- If any command cannot run due to environment limits, report exactly what was skipped and why.

## Output Contract

- List changed files.
- State why each file changed.
- Report validations run and outcomes.
- Report residual risks or follow-up checks, if any.
