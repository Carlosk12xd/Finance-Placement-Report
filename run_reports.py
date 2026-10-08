"""Optional headless entrypoint; the app uses exactly the same service."""
import argparse
from pathlib import Path
from finance.service import generate

parser = argparse.ArgumentParser(description="Generate the fixed-format Finance reports.")
for kind in ("students", "internships", "offers"):
    parser.add_argument(f"--{kind}", required=True)
    parser.add_argument(f"--{kind}-sheet")
    parser.add_argument(f"--{kind}-header-row", type=int, default=1)
parser.add_argument("--as-of", required=True, help="Actual export date, YYYY-MM-DD")
parser.add_argument("--previous-tracker")
parser.add_argument("--out-dir", required=True)
args = parser.parse_args()
inputs = {}
for kind in ("students", "internships", "offers"):
    path = Path(getattr(args, kind))
    inputs[kind] = {"data": path.read_bytes(), "name": path.name, "sheet": getattr(args, f"{kind}_sheet"), "header_row": getattr(args, f"{kind}_header_row")}
previous = None
if args.previous_tracker:
    path = Path(args.previous_tracker)
    previous = {"name": path.name, "data": path.read_bytes()}
result = generate(inputs, args.as_of, previous)
destination = Path(args.out_dir)
destination.mkdir(parents=True, exist_ok=True)
for name, data in result["files"].items():
    target = destination / name
    if target.exists():
        raise SystemExit(f"Output already exists: {target}. Use a new output directory.")
for name, data in result["files"].items():
    (destination / name).write_bytes(data)
print(f"Saved {len(result['files'])} files. Tracker: {result['manifest']['counts']['tracker']} students. Historical: {result['manifest']['counts']['historical']} students.")
