import argparse
import shutil
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.exc import PackageNotFoundError

from apply_theme_fonts import process_presentation
from define_theme_fonts import FontPolicy, load_font_policy, update_theme_fonts
from logger import Logger

__version__ = "2026-03-31"


def create_backup(path: Path) -> Path:
    backup_path = path.with_stem(f"{path.stem} - backup")
    backup_number = 2
    while backup_path.exists():
        backup_path = path.with_stem(f"{path.stem} - backup ({backup_number})")
        backup_number += 1
    shutil.copyfile(path, backup_path)
    return backup_path


def process_pptx_file(
    pptx_path: Path,
    preserve_code_fonts: bool,
    dry_run: bool = False,
    font_policy: FontPolicy | None = None,
) -> None:
    log_path = pptx_path.with_suffix(".log")
    with open(log_path, "a") as log_file:
        logger = Logger(log_file)

        if not dry_run:
            backup_path = create_backup(pptx_path)
            logger.log(f"{pptx_path} was backed up to {backup_path}.")

        presentation = Presentation(str(pptx_path))
        if dry_run:
            logger.log(f"{pptx_path} was opened. (dry run)")
        else:
            logger.log(f"{pptx_path} was opened.")

        if font_policy is not None:
            update_theme_fonts(presentation, font_policy, logger)

        process_presentation(presentation, preserve_code_fonts, logger)

        if not dry_run:
            presentation.save(str(pptx_path))
            logger.log(f"{pptx_path} was saved.")


def main() -> int:
    print(f"replace_fonts - version {__version__} by Shinichi Akiyama")

    parser = argparse.ArgumentParser(
        description="Replace fonts in PowerPoint presentations"
    )
    parser.add_argument(
        "files", nargs="*", metavar="FILE", help="PowerPoint (.pptx) files to process"
    )
    parser.add_argument("--code", help="preserve code fonts", action="store_true")
    parser.add_argument(
        "--dry-run",
        help="preview font replacements without modifying files",
        action="store_true",
    )
    parser.add_argument(
        "--font-policy",
        help="YAML file defining font policy for theme fonts",
        type=Path,
    )
    args = parser.parse_args()
    preserve_code_fonts = args.code
    dry_run = args.dry_run
    font_policy_path: Path | None = args.font_policy
    font_policy: FontPolicy | None = None
    if font_policy_path:
        try:
            font_policy = load_font_policy(font_policy_path)
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
            return 1

    if not args.files:
        print("No files specified.")
        return 0

    success_count = 0
    failure_count = 0

    for pptx_path_str in args.files:
        pptx_path = Path(pptx_path_str)
        try:
            process_pptx_file(pptx_path, preserve_code_fonts, dry_run, font_policy)
            success_count += 1
        except FileNotFoundError:
            print(f"Error: File not found: {pptx_path}")
            failure_count += 1
        except (PackageNotFoundError, zipfile.BadZipFile, KeyError):
            print(f"Error: Invalid PowerPoint file: {pptx_path}")
            failure_count += 1
        except Exception as e:
            print(f"Error processing {pptx_path}: {type(e).__name__}: {e}")
            failure_count += 1

    total = success_count + failure_count
    if failure_count > 0:
        print(
            f"Processing complete: {success_count} succeeded, "
            f"{failure_count} failed out of {total}."
        )
    else:
        print(f"All {total} file(s) processed successfully.")

    return 1 if failure_count > 0 else 0


if __name__ == "__main__":
    exit(main())
