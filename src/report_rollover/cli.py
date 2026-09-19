"""Command-line interface for the supported Excel report template."""
import argparse
from pathlib import Path
import sys

from .models import ValidationError
from .rollover import roll_batch


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Roll validated material reports into next month.")
    parser.add_argument("--input", required=True, type=Path, help="Directory of supported .xlsx reports")
    parser.add_argument("--output", required=True, type=Path, help="New directory outside the input directory")
    args = parser.parse_args(argv)
    try:
        destination = roll_batch(args.input, args.output)
    except ValidationError as exc:
        print(f"Validation error: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"File operation failed: {exc}", file=sys.stderr)
        return 1
    count = len(list(destination.glob("*.xlsx")))
    print(f"Created {count} report(s) and rollover-summary.csv in {destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
