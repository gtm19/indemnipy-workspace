# indemnipy

## Development

### Prerequisites

This project uses [uv](https://docs.astral.sh/uv/) for dependency and workspace management. Install it via:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### Workspace structure

This repo is a [uv workspace](https://docs.astral.sh/uv/concepts/workspaces/) — a single `uv.lock` at the root covers all packages under `packages/`. uv is workspace-aware, so most commands work identically whether you run them from the root or from inside an individual package directory:

```bash
uv run pytest                  # from root — runs tests across all packages
cd packages/indemnipy-ai
uv run pytest                  # from a package — runs only that package's tests
```

The shared lockfile means dependency resolution is always consistent across the workspace.

### Setup

Install all dependencies (including dev tools):

```bash
uv sync
```

### Pre-commit

[pre-commit](https://pre-commit.com/) hooks are configured to run on every commit. After `uv sync`, install the hooks with:

```bash
uv run pre-commit install
```

The hooks enforce:

- **ruff** — linting (with auto-fix) and formatting
- **uv-lock** — ensures `uv.lock` is up to date after any dependency changes
- End-of-file newlines, trailing whitespace, TOML validity, merge conflict markers, and consistent line endings (LF)

To run all hooks manually against the whole codebase:

```bash
uv run pre-commit run --all-files
```
