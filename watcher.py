"""
Document Review Watcher

A lightweight scheduler wrapper for the document review pipeline. Checks the
configured input folder for any new documents (.docx, .msg, or .eml) and runs
pipeline.py only if there is work to do. Designed to be invoked by a task
scheduler (e.g. Windows Task Scheduler, cron) so the pipeline doesn't run
needlessly when the inbox is empty.

License: MIT
"""

import os
import sys
import subprocess
import configparser
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
PIPELINE_SCRIPT = SCRIPT_DIR / "pipeline.py"
CONFIG_FILE = SCRIPT_DIR / "config.ini"
WATCHER_LOG = SCRIPT_DIR / "watcher.log"

SUPPORTED_EXTENSIONS = ('.docx', '.msg', '.eml')


def log(message: str):
    """Log a message to both console and watcher log file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"

    print(log_entry)

    with open(WATCHER_LOG, 'a', encoding='utf-8') as f:
        f.write(log_entry + "\n")


def get_watch_folder() -> str:
    """
    Read the input folder path from config.ini.

    Returns:
        The configured input folder path

    Raises:
        SystemExit: If config.ini is missing or input_folder is not set
    """
    if not CONFIG_FILE.exists():
        log(f"ERROR: config.ini not found at {CONFIG_FILE}")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)

    if 'Paths' not in config or 'input_folder' not in config['Paths']:
        log("ERROR: [Paths] input_folder not set in config.ini")
        sys.exit(1)

    return config['Paths']['input_folder']


def has_pending_documents(watch_folder: str) -> bool:
    """
    Check if there are any supported documents in the watch folder.

    Args:
        watch_folder: Path to the folder to check

    Returns:
        True if any .docx/.msg/.eml files are present, False otherwise
    """
    if not os.path.exists(watch_folder):
        log(f"ERROR: Watch folder not found: {watch_folder}")
        return False

    pending = [
        f for f in os.listdir(watch_folder)
        if f.lower().endswith(SUPPORTED_EXTENSIONS) and not f.startswith('~$')
    ]

    if pending:
        log(f"Found {len(pending)} document(s) to process")
        return True
    else:
        log("No documents found - nothing to process")
        return False


def run_pipeline() -> bool:
    """
    Run pipeline.py as a subprocess with a 20-minute timeout.

    Returns:
        True if pipeline completed successfully, False otherwise
    """
    try:
        log("Running pipeline...")

        process = subprocess.Popen(
            [sys.executable, str(PIPELINE_SCRIPT)],
            cwd=str(SCRIPT_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        try:
            stdout, stderr = process.communicate(timeout=1200)

            if process.returncode == 0:
                log("Pipeline completed successfully")
                if stdout:
                    log(f"Output: {stdout[:500]}")
                return True
            else:
                log(f"Pipeline failed with return code {process.returncode}")
                if stderr:
                    log(f"Error output: {stderr}")
                return False

        except subprocess.TimeoutExpired:
            process.kill()
            log("ERROR: Pipeline timed out after 20 minutes")
            return False

    except Exception as e:
        log(f"ERROR: Failed to run pipeline: {str(e)}")
        import traceback
        log(f"Traceback: {traceback.format_exc()}")
        return False


def main():
    """Main execution function."""
    log("=" * 60)
    log("Document Review Watcher - Starting")
    log("=" * 60)

    if not PIPELINE_SCRIPT.exists():
        log(f"ERROR: pipeline.py not found at {PIPELINE_SCRIPT}")
        sys.exit(1)

    watch_folder = get_watch_folder()

    if has_pending_documents(watch_folder):
        success = run_pipeline()

        if success:
            log("Watcher completed successfully")
            sys.exit(0)
        else:
            log("Watcher completed with errors")
            sys.exit(1)
    else:
        log("Watcher completed - no work to do")
        sys.exit(0)


if __name__ == "__main__":
    main()
