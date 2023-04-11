"""
Fuzzy find files in multiple paths on Windows using fzf.

- Files searched are listed in configfile (%LOCALAPPDATA%\finddoc\finddoc.toml)
- File lists are cached (use finddoc.py --update to refresh)

Dependencies:

    pip install pyperclip appdirs

    fzf from https://github.com/junegunn/fzf/releases

Config example (TOML format):

    [finddoc]
    paths = [
        "%USERPROFILE%\\Documents",
        "%ONEDRIVE%\\Documents",
        ...
    ]
"""

import argparse
import hashlib
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
from contextlib import contextmanager
from pathlib import Path

import appdirs
import pyperclip
import tomli
import tomli_w
import tqdm

FIELD_SEP = b"\t"
RECORD_SEP = b"\x00"

CACHE_DIR = Path(appdirs.user_cache_dir()) / "finddoc"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

IGNORE_RE = re.compile("\.(bkp|dtmp|part)$", re.IGNORECASE)


def compressvars(path):
    shortest_path = path
    for var, value in os.environ.items():
        if path.startswith(value):
            rest = path[len(value):]
            candidate = f'%{var}%{rest}'
            if len(candidate) < len(shortest_path):
                shortest_path = candidate
    return shortest_path


def find_totalcmd():
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        if program_files := os.environ.get(env):
            for exe in "totalcmd64.exe", "totalcmd.exe":
                path = Path(program_files) / "totalcmd" / exe
                if path.exists():
                    return path


explorer_exe = shutil.which("explorer")
totalcmd_exe = shutil.which("totalcmd64") or shutil.which(
    "totalcmd") or find_totalcmd()


def sanitize_text(text):
    """Sanitize a multiline text string for preview"""
    text = re.sub("([ \t\r+]*\n)", "\n", text)
    text = re.sub("\n{3,}", "\n\n", text)
    return text


def start_thread(target, args=None):
    """Create and start a thread"""
    if args is None:
        args = []
    thread = threading.Thread(target=target, args=args)
    thread.start()
    return thread


def parallel_walk(base):
    """Multithreaded version of os.walk"""
    jobs_created = 0

    def worker(jobs: queue.Queue, results: queue.Queue):
        nonlocal jobs_created
        while jobbase := jobs.get():
            dirs = []
            nondirs = []
            try:
                for entry in os.scandir(jobbase):
                    if entry.is_dir():
                        dirs.append(entry.name)
                        jobs_created += 1
                        jobs.put(jobbase / entry.name)
                    else:
                        nondirs.append(entry.name)
                results.put((jobbase, dirs, nondirs))
            except OSError:
                results.put(None)

    jobqueue = queue.Queue()
    resultqueue = queue.Queue()
    threads = [
        start_thread(worker, (jobqueue, resultqueue)) for _ in range(os.cpu_count())
    ]

    # start the job for the root
    jobs_created += 1
    jobqueue.put(Path(base))

    # wait for jobs to complete
    jobs_completed = 0
    while jobs_completed < jobs_created:
        result = resultqueue.get()
        if result:
            yield result
        jobs_completed += 1

    # stop worker threads
    for _ in threads:
        jobqueue.put(None)
    for thread in threads:
        thread.join()


def walk(base, dst, alt_dst=None, progress=None):
    """Walk directory and write paths to dst (and optionally alt_dst)"""
    for root, _dirs, files in parallel_walk(base):
        root = Path(root)
        if progress:
            progress.put(len(files))
        for file in files:
            fullpath = root / file
            str_path = str(fullpath)
            if IGNORE_RE.search(str_path):
                continue
            block = str_path.encode() + RECORD_SEP
            dst.write(block)
            if alt_dst:
                alt_dst.write(block)


def parse_path(path):
    """Replace environment variables in path and normalize"""
    path = os.path.expandvars(path)
    path = os.path.normpath(path)
    return path


def rescan(root, progress):
    digest = hashlib.sha256(root.encode()).hexdigest()
    list_path = CACHE_DIR / digest
    part_path = str(list_path) + ".part"
    with open(part_path, "wb") as outfile:
        walk(root, outfile, progress=progress)
    os.replace(part_path, list_path)


def update():
    progress_queue = queue.Queue()

    try:
        with open(CACHE_DIR / "files", "r") as infile:
            total_count = int(infile.read())
    except (FileNotFoundError, ValueError):
        total_count = 10000

    threads = [start_thread(rescan, (root, progress_queue)) for root in roots]

    def progress():
        total = 0
        with tqdm.tqdm(total=total_count, colour="blue", unit="files") as tq:
            total = 0
            while True:
                count = progress_queue.get()
                if count is None:
                    break
                total += count
                if total > tq.total:
                    tq.total = total
                tq.update(count)
                tq.set_description(f"{len(threading.enumerate())} threads")
        tq.set_description(f"{len(threading.enumerate())} threads")
        progress_queue.put(total)
        return total

    progress_thread = start_thread(progress)
    for thread in threads:
        thread.join()
    progress_queue.put(None)
    progress_thread.join()
    total_count = progress_queue.get()
    with open(CACHE_DIR / "files", "w") as outfile:
        outfile.write(str(total_count))


def cached_walk(root, io):
    """
    Walk directory tree and write nul-separated paths to `io`. Caches result in
    `CACHE_DIR`.
    """
    digest = hashlib.sha256(root.encode()).hexdigest()
    cache_path = CACHE_DIR / digest
    try:
        with open(cache_path, "rb") as infile:
            shutil.copyfileobj(infile, io)
    except FileNotFoundError:
        part = str(cache_path) + ".part"
        with open(part, "wb") as part_fileout:
            walk(root, io, part_fileout)
        os.replace(part, cache_path)


def fzf(opts):
    history_path = CACHE_DIR / "history"
    fzf_command = shutil.which("fzf")
    if not fzf_command:
        print(
            "fzf is needed and was not found path. Download from https://github.com/junegunn/fzf/releases"
        )
        exit(1)

    expect = "alt-u,alt-c,alt-e"
    if totalcmd_exe:
        expect = expect + ",alt-o"

    header = "enter=open, alt-c=copy path, alt-e=show in explorer, ctrl+p/n=history, alt-u=update, esc=abort"
    if totalcmd_exe:
        header += ", alt-o=show in totalcmd"

    command = [
        fzf_command,
        "--expect",
        expect,
        "--print0",
        "--read0",
        "--with-nth=1",
        "--delimiter=@",
        # "--border",
        "--history",
        history_path,
        "--header",
        header,
        "--bind",
        "shift-up:preview-page-up,shift-down:preview-page-down",
    ]
    if opts.preview:
        command.extend(
            [
                "--preview",
                f"{sys.executable} {__file__} preview {{}}",
                "--preview-window",
                "up,30%",
            ]
        )

    proc = subprocess.Popen(
        command, stdin=subprocess.PIPE, stdout=subprocess.PIPE)

    try:
        try:
            for root in roots:
                cached_walk(root, proc.stdin)
            proc.stdin.close()
        except (BrokenPipeError, OSError):
            pass
        if output := proc.stdout.read():
            key, path, _ = output.split(RECORD_SEP)
            if key == b"":
                os.startfile(path.decode())
            elif key == b"alt-c":
                pyperclip.copy(path.decode())
            elif key == b"alt-e":
                os.system(f'explorer.exe /select,"{path.decode()}"')
            elif key == b"alt-o":
                subprocess.call((totalcmd_exe, "/a", "/o", path.decode()))
            elif key == b"alt-u":
                update()
                return fzf(opts)

    finally:
        try:
            proc.stdin.close()
        except OSError:
            pass
        proc.stdout.close()


def preview(path):
    """Preview a file. Not very good."""
    _path, ext = os.path.splitext(path)
    if ext in (".doc",):
        command = shutil.which("catdoc")
        if command:
            output = subprocess.check_output(
                (command, "-s", "8859-1", path), cwd=os.path.dirname(command)
            )
            print(sanitize_text(output.decode(encoding="iso-8859-1")))
        else:
            print(".doc preview requires catdoc.exe")
    elif ext in (".doc", ".docx"):
        import win32com.client

        word = win32com.client.Dispatch("Word.Application")
        word.visible = False
        word.Documents.Open(path)
        doc = word.ActiveDocument
        text = (
            doc.Range().Text.replace("\x01", "").replace("\x07", "").replace("\r", "\n")
        )
        text = re.sub("([ \t+]*\n){3,}", "\n", text)
        doc.Close()
        print(text.strip())
    else:
        print("Unknown file format")


@contextmanager
def edit_config():
    with open(opts.config, "rb") as infile:
        config = tomli.load(infile)
    yield config
    part_config = opts.config + ".part"
    with open(part_config, "wb") as outfile:
        tomli_w.dump(config, outfile)
    os.replace(part_config, opts.config)


def listdirs():
    """List configured directories"""
    for root in roots:
        print(root)


def add(path):
    absolute_path = Path(path).absolute()
    add_path = compressvars(str(absolute_path))
    with edit_config() as config:
        config["finddoc"]["paths"].append(add_path)
        print(f"Added '{add_path}' to list")


def remove(path):
    absolute_path = Path(path).absolute()
    findpath = compressvars(str(absolute_path))

    def keep_path(path):
        is_keeeper = Path(parse_path(path)).absolute() != absolute_path
        if not is_keeeper:
            print(f"Removed '{findpath}' from list")
        return is_keeeper

    with edit_config() as config:
        config["finddoc"]["paths"] = list(
            filter(keep_path, config["finddoc"]["paths"]))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    default_config = os.path.join(
        appdirs.user_config_dir(), "finddoc", "finddoc.toml")
    parser.add_argument("--preview", action="store_const",
                        const=True, default=False)
    parser.add_argument("--config", default=default_config,
                        help=f"Path to config file (default: {default_config})")

    subs = parser.add_subparsers()
    subs.dest = "command"
    subs.default = "find"
    preview_parser = subs.add_parser("preview", help="Preview file")
    preview_parser.add_argument("file")
    update_parser = subs.add_parser("update", help="Update directory caches")
    find_parser = subs.add_parser("find", help="Find files (default)")
    list_parser = subs.add_parser("list", help="List included directories")

    add_parser = subs.add_parser("add", help="Add path to search tree")
    add_parser.add_argument("path", help="Path to scan for files")

    remove_parser = subs.add_parser("remove", help="Remove path from search")
    remove_parser.add_argument("path", help="Path to remove")

    opts = parser.parse_args()

    with open(opts.config, "rb") as infile:
        config = tomli.load(infile)
    roots = [parse_path(path) for path in config["finddoc"]["paths"]]

    if opts.command == "list":
        listdirs()
    elif opts.command == "find":
        fzf(opts)
    elif opts.command == "preview":
        preview(opts.file)
    elif opts.command == "update":
        update()
    elif opts.command == "add":
        add(opts.path)
    elif opts.command == "remove":
        remove(opts.path)
