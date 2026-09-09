"""PDF conversion utility using LibreOffice headless mode."""

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple


def find_libreoffice_binary() -> Optional[str]:
    """Finds the LibreOffice executable on macOS, Linux, or Windows."""
    # Check PATH first
    for candidate in ["soffice", "libreoffice"]:
        found = shutil.which(candidate)
        if found:
            return found

    # Check common absolute paths
    common_paths = [
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Linux
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/usr/local/bin/libreoffice",
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    for p in common_paths:
        if os.path.isfile(p) and os.access(p, os.X_OK):
            return p

    return None


def convert_docx_to_pdf(
    docx_path: str | Path,
    output_dir: Optional[str | Path] = None,
    timeout_sec: int = 60
) -> Tuple[Optional[Path], Optional[str]]:
    """
    Converts a DOCX file to PDF using LibreOffice headless mode.
    Returns (pdf_path, error_message).
    If LibreOffice is not installed, returns (None, help_text).
    """
    docx_path = Path(docx_path).resolve()
    if not docx_path.is_file():
        return None, f"Source DOCX file not found: {docx_path}"

    target_dir = Path(output_dir).resolve() if output_dir else docx_path.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    soffice = find_libreoffice_binary()
    if not soffice:
        # On macOS, check if Apple Pages is available as a native fallback
        if sys.platform == "darwin" and os.path.isdir("/Applications/Pages.app"):
            expected_pdf = target_dir / f"{docx_path.stem}.pdf"
            script = f'''
            tell application "Pages"
                set myDoc to open POSIX file "{docx_path.resolve().as_posix()}"
                export myDoc to POSIX file "{expected_pdf.resolve().as_posix()}" as PDF
                close myDoc saving no
            end tell
            '''
            try:
                res = subprocess.run(
                    ["osascript", "-e", script],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=timeout_sec,
                )
                # Quit Pages in background if no open documents
                subprocess.run(
                    ["osascript", "-e", 'tell application "Pages" to if (count of documents) = 0 then quit'],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                if expected_pdf.is_file():
                    return expected_pdf, None
            except Exception:
                pass

        msg = (
            "LibreOffice is not installed on this server.\n"
            "To enable PDF generation, install LibreOffice:\n"
            "• macOS: brew install --cask libreoffice\n"
            "• Linux/Ubuntu: sudo apt-get install libreoffice\n"
            "• Windows: install LibreOffice from https://www.libreoffice.org"
        )
        return None, msg

    temp_profile_dir = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    try:
        # file:// URL format for LibreOffice UserInstallation
        profile_url = f"file://{temp_profile_dir.resolve().as_posix()}"
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_url}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(target_dir),
            str(docx_path),
        ]

        res = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout_sec,
            check=False,
        )
        expected_pdf = target_dir / f"{docx_path.stem}.pdf"
        if expected_pdf.is_file():
            return expected_pdf, None
        return None, f"LibreOffice exited with code {res.returncode}. Stderr: {res.stderr.strip()}"
    except subprocess.TimeoutExpired:
        return None, f"LibreOffice conversion timed out after {timeout_sec} seconds."
    except Exception as exc:
        return None, f"Failed to run LibreOffice: {exc}"
    finally:
        shutil.rmtree(temp_profile_dir, ignore_errors=True)
