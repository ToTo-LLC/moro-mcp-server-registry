# Moro MCP Server Registry

Pre-configured MCP server templates for the Moro AI Chat platform.

## Structure

- `registry.json` — Index of all available server templates
- `templates/<id>.json` — Individual server template files
- `icons/<id>.svg` — Server icons (SVG preferred, PNG accepted)

## Template Types

- **full** — Connection-ready with auth config, setup instructions, and credentials
- **stub** — Minimal metadata; relies on Smart Discovery to fill connection details

## Template Schema

### registry.json entry

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| id | string | yes | Unique server identifier (kebab-case) |
| name | string | yes | Display name |
| icon_url | string | no | Relative path to icon (e.g., `icons/slack.svg`) |
| category | string | no | Category for grouping |
| url_pattern | string | no | fnmatch glob for URL matching (e.g., `*.slack.com`) |
| template_path | string | yes | Relative path to template JSON |
| template_type | string | yes | `full` or `stub` |

### Template JSON

| Field | Type | Description |
|-------|------|-------------|
| id | string | Must match registry entry id |
| name | string | Display name |
| description | string | What this server provides |
| category | string | Category |
| icon_url | string | Relative path to icon |
| server_type | string\|null | `stdio`, `sse`, or `http` |
| command | string\|null | Command for stdio (e.g., `npx`) |
| args | string[]\|null | Command arguments |
| url | string\|null | URL for sse/http |
| org_auth_type | string\|null | `api_key`, `oauth2`, `bearer`, `none` |
| user_auth_type | string\|null | `api_key`, `oauth2`, `bearer`, `none` |
| oauth_config | object\|null | OAuth endpoints and scopes |
| credentials_needed | string[]\|null | Required credential keys |
| user_auth_instructions | string\|null | End-user auth setup steps |
| setup_notes | string\|null | Admin setup instructions |
| env_mapping | object\|null | Environment variable mapping |
| tags | string[] | Search tags |
| source | object\|null | Provenance: package, repository, docs URLs |
