import tempfile
from pathlib import Path

import nox

# set backend to uv or virtualenv
nox.options.default_venv_backend = "uv|virtualenv"

PACKAGE_PATH = Path(__file__).parent / "packages"


def install_all_packages(session):
    """Install all packages in the workspace."""
    for package_dir in PACKAGE_PATH.iterdir():
        if package_dir.is_dir() and (package_dir / "pyproject.toml").exists():
            session.install("-e", str(package_dir))


@nox.session(python=["3.10", "3.12", "3.13"])
def tests(session):
    """Run the test suite."""
    install_all_packages(session)
    session.install("--group", "test")
    session.run("pytest")


@nox.session(python=["3.10", "3.12", "3.13"], default=False)
def smoke(session):
    """Build indemnipy-ai from source and verify the getting-started imports work."""
    with tempfile.TemporaryDirectory() as dist_dir:
        session.run(
            "uv",
            "build",
            "--package",
            "indemnipy-ai",
            "--out-dir",
            dist_dir,
            external=True,
        )
        wheels = list(Path(dist_dir).glob("indemnipy_ai-*.whl"))
        session.install(str(wheels[0]))

    session.run("python", "scripts/smoke.py")


@nox.session(default=False)
def build_docs(session):
    """Build the documentation."""
    session.install("--group", "docs")
    session.run("zensical", "build", "--clean")


@nox.session(default=False)
def lint(session):
    """Run linting and formatting checks."""
    session.install("--group", "lint")
    session.run("ruff", "format", ".")
    session.run("ruff", "check", "--fix", ".")


@nox.session(default=False)
def typecheck(session):
    """Run type checking."""
    session.install("--group", "typecheck", "--group", "test")
    src_dirs = [str(p) for p in Path("packages").glob("*/src")]
    session.run("basedpyright", *src_dirs, "--level", "error")
