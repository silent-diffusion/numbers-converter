"""The numbers-converter command: saves Excel copies of Numbers files, and of the Numbers files in
folders, without the app's window, for batches and scripts. It runs in the app's helper
(numbers_tool.py started without --serve), which hands it what it needs as `tools` (CLI_TOOLS there).
find() also finds the Numbers files in folders opened in the app.

  numbers-converter Budget.numbers
  numbers-converter "C:\\Reports" --out "D:\\Excel" --json > summary.json
"""

import argparse
import getpass
import glob
import json
import os
import re
import stat
import sys

EXIT_OK, EXIT_FAILED, EXIT_USAGE = 0, 1, 2

# Folders with none of the user's files in them: Macs add __MACOSX to zip files, Windows the others to drives.
SKIP_FOLDERS = {"__macosx", "$recycle.bin", "system volume information"}
LINK_TAGS = {0xA0000003, 0xA000000C}  # junctions and symbolic links, which can lead back up the tree

DESCRIPTION = """\
Saves an Excel (.xlsx) copy of each Numbers file, on this PC: nothing is uploaded. Folders are
searched, with their subfolders, for the Numbers files in them, and wildcards work: *.numbers"""

EPILOG = """\
exit codes:
  0  every file was saved (with --dry-run: can be)
  1  a file couldn't be converted, or needs a password
  2  no Numbers files were found, or the command has a mistake in it

examples:
  numbers-converter Budget.numbers
  numbers-converter "C:\\Reports" --out "D:\\Excel"
  numbers-converter *.numbers --layout one --json > summary.json"""

LAYOUTS = {"separate": "tables", "one": "sheet"}


# ---------------------------------------------------------------- finding Numbers files


def is_numbers(path):
    return path.rstrip("\\/").lower().endswith(".numbers")


def natural(name):
    """Sorts names as Explorer does: "Q2" before "Q10", whatever their case."""
    return [int(part) if part.isdigit() else part.casefold() for part in re.split(r"(\d+)", name)]


def is_link(entry):
    try:
        if entry.is_symlink():
            return True
        if not entry.stat(follow_symlinks=False).st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT:
            return False
        return os.lstat(entry.path).st_reparse_tag in LINK_TAGS  # not OneDrive's files and folders, which are too
    except AttributeError:  # not on Windows
        return False
    except OSError:  # can't tell where it leads, so it isn't followed
        return True


def find(paths, limit=None, max_folders=None):
    """The Numbers files in `paths`: .numbers files and Mac package folders as they are, and those in other
    folders and their subfolders (in name order, a folder's own files before its subfolders'). Returns
    {"files": [{"path", "folder"}], "skipped", "missing", "empty", "more"}: "folder" is the folder a file
    was found in (None for one given itself); "skipped" lists paths that aren't Numbers files or folders,
    "missing" those that don't exist, and "empty" folders with no Numbers files; "more" says the search
    stopped at `limit` files or after searching `max_folders` folders."""
    found = {"files": [], "skipped": [], "missing": [], "empty": [], "more": False}
    seen = set()
    searched = 0

    def add(path, folder):  # False once there are `limit` files
        key = os.path.normcase(os.path.abspath(path))
        if key not in seen:
            seen.add(key)
            found["files"].append({"path": path, "folder": folder})
        return limit is None or len(found["files"]) < limit

    for path in paths:
        if found["more"]:
            break
        if not os.path.exists(path):
            found["missing"].append(path)
        elif is_numbers(path):
            found["more"] = not add(path, None)
        elif not os.path.isdir(path):
            found["skipped"].append(path)
        else:
            before = len(found["files"])
            folders = [path]
            while folders and not found["more"]:
                if max_folders is not None and searched >= max_folders:
                    found["more"] = True
                    break
                searched += 1
                try:
                    with os.scandir(folders.pop()) as listing:
                        entries = sorted(listing, key=lambda e: natural(e.name))
                except OSError:  # e.g. a folder this user can't open
                    continue
                subfolders = []
                for entry in entries:
                    # Hidden files are left out, and with them the "._Budget.numbers" files Macs leave on drives.
                    if entry.name.startswith(".") or entry.name.lower() in SKIP_FOLDERS or is_link(entry):
                        continue
                    if is_numbers(entry.name):
                        if not add(entry.path, path):
                            found["more"] = True
                            break
                    else:
                        try:
                            if entry.is_dir():
                                subfolders.append(entry.path)
                        except OSError:
                            pass
                folders.extend(reversed(subfolders))  # popped from the end, so searched in name order
            if len(found["files"]) == before and not found["more"]:
                found["empty"].append(path)
    return found


def expand(args):
    """The paths the command was given, with wildcards (which Windows' shells leave as they are) expanded:
    (paths, patterns that matched nothing). What a wildcard matches that isn't a Numbers file or a
    folder is left out, as asking for * means every Numbers file there."""
    paths, unmatched = [], []
    for arg in args:
        if not ("*" in arg or "?" in arg) or os.path.exists(arg):  # * and ? can't be in a Windows file name
            paths.append(arg)
            continue
        matches = [m for m in glob.glob(arg.replace("[", "[[]")) if is_numbers(m) or os.path.isdir(m)]
        if matches:
            paths += sorted(matches, key=lambda m: natural(os.path.basename(m.rstrip("\\/"))))
        else:
            unmatched.append(arg)
    return paths, unmatched


# ---------------------------------------------------------------- the command


def parse(argv, version):
    parser = argparse.ArgumentParser(prog="numbers-converter", description=DESCRIPTION, epilog=EPILOG,
                                     formatter_class=argparse.RawDescriptionHelpFormatter, add_help=False)
    parser.add_argument("paths", nargs="+", metavar="FILE_OR_FOLDER",
                        help="Numbers files, Mac package folders, and folders to search for Numbers files")
    parser.add_argument("-o", "--out", metavar="FOLDER",
                        help="save the Excel files in FOLDER (in the subfolders they were found in) instead of next "
                             "to the originals")
    parser.add_argument("--layout", choices=list(LAYOUTS), default="separate",
                        help="sheets with more than one table: each table on an Excel sheet of its own (separate, the "
                             "default), or all of them on one sheet, placed as in Numbers (one)")
    parser.add_argument("--contents", action="store_true", help="start each Excel file with a sheet linking to every table")
    parser.add_argument("--overwrite", action="store_true",
                        help="write over Excel files that are already there, instead of adding \" (1)\" to the name")
    parser.add_argument("--password-file", metavar="FILE",
                        help="try the password on the first line of FILE on locked files (in a terminal, you're asked "
                             "for any password that's needed)")
    parser.add_argument("-n", "--dry-run", action="store_true",
                        help="say what would be saved, and what changes in Excel, without saving anything")
    parser.add_argument("--json", action="store_true", help="print a summary in JSON, for scripts")
    parser.add_argument("--version", action="version", version=f"numbers-converter {version}", help="show the version")
    parser.add_argument("-h", "--help", action="help", help="show this help")
    args = parser.parse_args(argv)
    if args.out and os.path.isfile(args.out):
        parser.error(f"--out needs a folder, and {args.out} is a file")
    if args.password_file:
        try:
            with open(args.password_file, encoding="utf-8-sig") as f:
                args.password = f.readline().rstrip("\r\n")
        except OSError as err:
            parser.error(f"can't read the password file: {err.strerror}")
    else:
        args.password = None
    return args


def shown(path):
    """A path as the command prints it: from the current folder when it's in there, otherwise in full."""
    full = os.path.abspath(path)
    try:
        relative = os.path.relpath(full)
    except ValueError:  # on another drive
        return full
    return full if relative.startswith("..") else relative


def can_type():
    """Whether someone can type a password: input comes from a console, not a file, a program or NUL
    (which Windows says is a terminal too)."""
    if not sys.stdin.isatty():
        return False
    if os.name != "nt":
        return True
    import ctypes
    import msvcrt

    mode = ctypes.c_uint32()
    return bool(ctypes.windll.kernel32.GetConsoleMode(msvcrt.get_osfhandle(sys.stdin.fileno()), ctypes.byref(mode)))


def counted(n, one, many=None):
    return f"{n:,} {one if n == 1 else many or one + 's'}"


class Run:
    """One run of the command: where each Excel copy goes, the passwords to try, and the results."""

    def __init__(self, args, tools):
        self.args = args
        self.tools = tools
        self.layout = {"default": LAYOUTS[args.layout]}
        self.password = args.password  # from --password-file, then the last one that worked
        self.ask = can_type()
        self.taken = set()  # Excel copies this run saves (or would), which it never writes over
        self.results = []

    def target(self, path, folder):
        """Where the Excel copy of `path` goes: next to it, or in --out (in the subfolder of `folder`, the
        folder given, that it was found in). Never over another file, except with --overwrite, and never
        over a copy this run saved."""
        stem = re.sub(r"\.numbers$", "", os.path.basename(path.rstrip("\\/")), flags=re.I)
        where = os.path.dirname(os.path.abspath(path.rstrip("\\/")))
        if self.args.out:
            sub = os.path.relpath(where, os.path.abspath(folder)) if folder else "."
            where = os.path.normpath(os.path.join(os.path.abspath(self.args.out), sub))
        out, n = os.path.join(where, stem + ".xlsx"), 0
        while os.path.normcase(out) in self.taken or (not self.args.overwrite and os.path.exists(out)):
            n += 1
            out = os.path.join(where, f"{stem} ({n}).xlsx")
        self.taken.add(os.path.normcase(out))
        return out

    def save(self, path, out, password):
        if self.args.dry_run:
            report = self.tools.check(path, password, self.layout)
            return {"status": "ready", "output": out, "changes": report["changes"], "notes": report["notes"]}
        os.makedirs(os.path.dirname(out), exist_ok=True)
        stats = self.tools.convert(path, out, password, self.layout, self.args.contents)
        report = self.tools.report(path, password, self.layout)
        kept = {k: stats[k] for k in ("sheets", "formulas", "formulasAsValues", "charts", "drawings", "pivots") if k in stats}
        return {"status": "saved", "output": out, **kept, "changes": report["changes"], "notes": report["notes"]}

    def convert(self, item):
        path = item["path"]
        result = {"input": os.path.abspath(path)}
        out = self.target(path, item["folder"])
        password, tried = None, []
        while True:
            try:
                result.update(self.save(path, out, password))
                if password:
                    self.password = password
                break
            except self.tools.UserError as err:
                if not err.extra.get("needsPassword"):
                    result.update(status="failed", error=str(err),
                                  **{k: err.extra[k] for k in ("reason", "details") if err.extra.get(k)})
                    break
                hint = err.extra.get("hint")
                if self.password and self.password not in tried:
                    password = self.password
                elif self.ask and len(tried) < 3 + bool(self.password):
                    password = self.prompt(path, hint, wrong=bool(tried))
                else:
                    password = None
                if not password:
                    result.update(status="locked", error=str(err), **({"hint": hint} if hint else {}))
                    break
                tried.append(password)
            except OSError as err:  # e.g. --out can't be made
                result.update(status="failed", reason="save", error=f"Couldn’t save {shown(out)}: {err.strerror or err}")
                break
            except Exception as err:  # never let one file stop the rest
                result.update(status="failed", reason="failed", error="This file couldn’t be converted.",
                              details=self.tools.technical(err, path))
                break
        if result["status"] not in ("saved", "ready"):
            self.taken.discard(os.path.normcase(out))
        self.results.append(result)
        return result

    @staticmethod
    def prompt(path, hint, wrong):
        lead = "That password isn’t right. " if wrong else ""
        hint = f" (hint: {hint})" if hint else ""
        try:
            return getpass.getpass(f"{lead}Password for {shown(path)}{hint}, or Enter to skip: ")
        except EOFError:
            return ""


# How a file's result is printed: its label, then what it says about it.
LABELS = {"saved": "Saved", "ready": "Would save", "locked": "Locked", "failed": "Failed"}


def describe(result):
    status = result["status"]
    if status in ("saved", "ready"):
        text = shown(result["output"])
        changes = result["changes"]
        if changes:
            text += f" — {counted(len(changes), 'change')} to check: " + "; ".join(c["title"] for c in changes)
        elif status == "ready":
            text += " — converts fully"
        return text
    text = f"{shown(result['input'])} — {result['error']}"
    if result.get("hint"):
        text += f" Hint: {result['hint']}"
    if result.get("details"):
        text += f" ({result['details']})"
    return text


def summary(results, dry_run):
    tally = {status: sum(r["status"] == status for r in results) for status in LABELS}
    parts = [counted(tally["ready"], "file") + " ready to save" if dry_run else f"{tally['saved']:,} saved"]
    parts += [f"{tally[s]:,} {s}" for s in ("locked", "failed") if tally[s]]
    return ", ".join(parts) + "." + (" Nothing was saved." if dry_run else "")


def configure_output():
    """Unicode on the console as it is; UTF-8 when redirected to a file or another program (where Python
    would otherwise use the ANSI code page and fail on names it can't spell)."""
    for stream in (sys.stdout, sys.stderr):
        if not stream.isatty() and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def main(argv, tools):
    configure_output()
    args = parse(argv, tools.version)
    paths, unmatched = expand(args.paths)
    found = find(paths)
    run = Run(args, tools)
    for reason, error, listed in (("missing", "This file doesn’t exist.", found["missing"]),
                                  ("not-numbers", "This isn’t a Numbers file.", found["skipped"])):
        run.results += [{"input": os.path.abspath(p), "status": "failed", "reason": reason, "error": error} for p in listed]
    warnings = [f"Nothing matches {p}" for p in unmatched] + [f"No Numbers files in {shown(p)}" for p in found["empty"]]
    if not found["files"] and not run.results and not warnings:
        warnings.append("No Numbers files found")

    pad = max(len(LABELS[s]) for s in ("ready" if args.dry_run else "saved", "locked", "failed")) + 2

    def say(result):
        if not args.json:
            print(f"{LABELS[result['status']]:<{pad}}{describe(result)}", flush=True)

    if not args.json:
        for warning in warnings:
            print(warning + ".", file=sys.stderr)
        for result in run.results:
            say(result)
        if len(found["files"]) > 1:
            print(f"{'Checking' if args.dry_run else 'Converting'} {counted(len(found['files']), 'Numbers file')}"
                  f"{' (nothing is saved)' if args.dry_run else ''}…", flush=True)
    stopped = False
    try:
        for item in found["files"]:
            say(run.convert(item))
    except KeyboardInterrupt:
        stopped = True

    results = run.results
    if args.json:
        tally = {status: sum(r["status"] == status for r in results) for status in LABELS}
        print(json.dumps({"version": tools.version, "dryRun": args.dry_run, **tally, "stopped": stopped,
                          "warnings": warnings, "files": results}, indent=2))
    elif stopped or len(results) > 1:
        print(("Stopped. " if stopped else "") + summary(results, args.dry_run))
    if not results:
        return EXIT_USAGE
    return EXIT_OK if not stopped and all(r["status"] in ("saved", "ready") for r in results) else EXIT_FAILED
