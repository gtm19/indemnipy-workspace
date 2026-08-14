import nox

# set backend to uv or virtualenv
nox.options.default_venv_backend = "uv|virtualenv"


@nox.session(default=False)
def build_docs(session):
    """Build the documentation."""
    session.install("--group", "docs")
    session.run("zensical", "build", "--clean")
