from pathlib import Path
import py_compile
import sqlite3

PROJECT = Path(__file__).resolve().parents[1]
APP = PROJECT / "app" / "app.py"
DB = PROJECT / "database" / "bam_motor_group.db"


def test_project_folders_exist():
    for name in ["app", "database", "uploads", "reports", "backups", "docs", "tests", "tools"]:
        assert (PROJECT / name).exists(), f"Missing folder: {name}"


def test_application_compiles():
    py_compile.compile(str(APP), doraise=True)


def test_database_opens():
    connection = sqlite3.connect(DB)
    connection.execute("PRAGMA integrity_check").fetchone()
    connection.close()


if __name__ == "__main__":
    test_project_folders_exist()
    test_application_compiles()
    test_database_opens()
    print("BAM Dealer Enterprise v16.0 health check passed.")
