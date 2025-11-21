# Repository Guidelines

## Project Structure & Module Organization
- Entry point: `edm_wizard.py` launches the PyQt5 wizard.
- Core modules: `edm_wizard/ui/pages/*.py` for each step; `ui/components/` for shared widgets; `utils/` for constants, data transforms, and XML generation; `workers/threads.py` for background tasks; `api/pas_client.py` for PAS lookups.
- Docs live in `README/` (Quickstart, refactoring guides); sample Access DB `ELESdxdb.mdb` and supporting scripts sit in the repo root.

## Build, Test, and Development Commands
- `pip install -r requirements_wizard.txt` — install pandas, SQLAlchemy, PyQt5, pyodbc, etc.
- `python edm_wizard.py` — run the wizard locally (preferred for dev).
- `build_exe.bat` — build the Windows executable with PyInstaller; outputs to `dist/`.
- Ensure the Microsoft Access Database Engine is installed to read `.mdb/.accdb` files.

## Coding Style & Naming Conventions
- Follow PEP 8: 4-space indent, snake_case for functions/vars, UPPER_SNAKE_CASE for constants (`utils/constants.py`), PascalCase for classes/pages.
- Keep UI page modules named `*_page.py`; align classes to match file names.
- Add type hints/docstrings on public functions; keep UI logic thin and push data transforms into `utils` or worker threads.

## Testing Guidelines
- No formal test suite yet; run manual smoke checks:
  - `python edm_wizard.py` with `ELESdxdb.mdb` to export to Excel.
  - Load a multi-sheet Excel, map MFG/MFG PN/Description, enable combine, and generate XML; confirm `{filename}_MFG.xml` and `{filename}_MFGPN.xml` exist.
  - Validate PAS search flows if you change `api/pas_client.py`.
- Before shipping an exe, run the same flows from `dist/EDM_Library_Wizard.exe`.

## Commit & Pull Request Guidelines
- Commit messages: concise, imperative summaries (see `git log`, e.g., "Add bulk clear MFG ...").
- PRs should include: what changed, why, screenshots/GIFs of UI changes, steps to reproduce, and data files used for manual testing. Link related issues/tickets when available.

## Security & Configuration Tips
- Do not commit customer or generated data (`*.xlsx`, `*.xml`, `dist/`); keep sensitive config/credentials out of the repo.
- Prefer environment variables or local config files for any API keys used by PAS; never hardcode secrets.
