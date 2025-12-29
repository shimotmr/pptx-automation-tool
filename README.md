```markdown
# pptx-automation-tool
pptx-automation-tool

## Quick Start

- **Create venv**: `python3 -m venv .venv`
- **Install deps**: `./.venv/bin/pip install -r requirements.txt` or `make install`
- **Run**: `./.venv/bin/python app.py` or `make run` or `scripts/run.sh`

This repository contains the main program `app.py` and `ppt_processor.py`.
Update `requirements.txt` with any dependencies before running `make install`.

**Files of interest**:
- `app.py`: entrypoint
- `ppt_processor.py`: PPTX processing helpers
- `requirements.txt`: Python dependencies
- `Makefile`: convenience targets (`install`, `run`, `test`)
- `pyproject.toml`: project metadata

```
# pptx-automation-tool
pptx-automation-tool
