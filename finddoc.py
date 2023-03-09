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
import tomli

import pyperclip
import appdirs

FIELD_SEP = b"\t"
RECORD_SEP = b"\x00"

CACHE_DIR = os.path.join(appdirs.user_cache_dir(), "finddoc")
os.makedirs(CACHE_DIR, exist_ok=True)

IGNORE_RE = re.compile("\.(bkp|dtmp|part)$", re.IGNORECASE)

def find_totalcmd():
    for env in ('ProgramFiles', 'ProgramFiles(x86)'):
        if program_files := os.environ.get(env):
            for exe in 'totalcmd64.exe', 'totalcmd.exe':
                path = os.path.join(
                    program_files, 'totalcmd', 'totalcmd64.exe')
                if os.path.exists(path):
                    return path


explorer_exe = shutil.which('explorer')
totalcmd_exe = shutil.which('totalcmd64') or shutil.which(
    'totalcmd') or find_totalcmd()


def sanitize_text(text):
    """Sanitize a multiline text string for preview"""
    text = re.sub("([ \t\r+]*\n)", "\n", text)
    text = re.sub("\n{3,}", "\n\n", text)
    return text


def make_thread(target, args):
    """Create and start a thread"""
    thread = threading.Thread(target=target, args=args)
    thread.start()
    return thread


def parwalk(base):
    """Multithreaded version of os.walk"""
    jobs_created = [0]
    jobs_completed = [0]

    def worker(jobs: queue.Queue, results: queue.Queue):
        while jobbase := jobs.get():
            dirs = []
            nondirs = []
            try:
                for entry in os.scandir(jobbase):
                    if entry.is_dir():
                        dirs.append(entry.name)
                        jobs_created[0] += 1
                        jobs.put(os.path.join(jobbase, entry.name))
                    else:
                        nondirs.append(entry.name)
                results.put((jobbase, dirs, nondirs))
            except OSError:
                results.put(None)

    jobqueue = queue.Queue()
    resultqueue = queue.Queue()
    threads = [
        make_thread(worker, (jobqueue, resultqueue)) for _ in range(os.cpu_count())
    ]

    jobs_created = [1]
    jobqueue.put(base)
    while jobs_completed[0] < jobs_created[0]:
        result = resultqueue.get()
        if result:
            yield result
        jobs_completed[0] += 1

    for _ in threads:
        jobqueue.put(None)
    for thread in threads:
        thread.join()


def walk(base, dst, alt_dst=None):
    """Walk directory and write paths to dst (and optionally alt_dst)"""
    for root, _dirs, files in parwalk(base):
        for file in files:
            fullpath = os.path.join(root, file)
            if IGNORE_RE.search(fullpath):
                continue
            block = fullpath.encode() + RECORD_SEP
            dst.write(block)
            if alt_dst:
                alt_dst.write(block)


def parse_path(path):
    """Replace environment variables in path and normalize"""
    path = os.path.expandvars(path)
    path = os.path.normpath(path)
    return path


def scan(root):
    digest = hashlib.sha256(root.encode()).hexdigest()
    list_path = os.path.join(CACHE_DIR, digest)
    part_path = list_path + ".part"
    with open(part_path, "wb") as outfile:
        walk(root, outfile)
    os.replace(part_path, list_path)


def start_thread(root):
    return make_thread(scan, (root,))


def update():
    threads = [start_thread(root) for root in roots]
    for thread in threads:
        thread.join()


def tee(infile, outfile1, outfile2):
    while data := infile.read(65536):
        outfile1.write(data)
        outfile2.write(data)


def fzf(opts):
    history_path = os.path.join(CACHE_DIR, 'history')
    fzf = shutil.which('fzf')
    if not fzf:
        print("fzf is needed and was not found path. Download from https://github.com/junegunn/fzf/releases")
        exit(1)

    expect = 'alt-c,alt-e'
    if totalcmd_exe:
        expect = expect + ',alt-o'

    command = [
        fzf,
        "--expect",
        expect,
        "--print0",
        "--read0",
        "--with-nth=1",
        "--delimiter=@",
        "--border",
        "--history",
        history_path,
        "--header",
        "enter=open file, alt-c=copy, alt-e=goto, ctrl+p/n=history, esc=abort",
        "--bind",
        "shift-up:preview-page-up,shift-down:preview-page-down",
    ]
    if opts.preview:
        command.extend([
            "--preview",
            f"{sys.executable} {__file__} preview {{}}",
            "--preview-window",
            "up,30%"
        ])

    proc = subprocess.Popen(command,
                            stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE,
                            )

    try:
        try:
            for root in roots:
                digest = hashlib.sha256(root.encode()).hexdigest()
                list_path = os.path.join(CACHE_DIR, digest)
                if os.path.exists(list_path):
                    with open(list_path, "rb") as infile:
                        shutil.copyfileobj(infile, proc.stdin)
                else:
                    part_file = list_path + ".part"
                    with open(part_file, "wb") as part_fileout:
                        walk(root, proc.stdin, part_fileout)
                    os.replace(part_file, list_path)
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
                # print(f'explorer.exe /select,"{path.decode()}"')
                # subprocess.call(("explorer.exe", f"/select,\"{path.decode()}\""))
                os.system(f'cmd.exe /c explorer.exe /select,"{path.decode()}"')
            elif key == b"alt-o":
                subprocess.call((totalcmd_exe, "/a", "/o", path.decode()))

    finally:
        try:
            proc.stdin.close()
        except OSError:
            pass
        proc.stdout.close()


def preview(path):
    """Preview a file. Not very good."""
    _path, ext = os.path.splitext(path)
    if ext in ('.doc',):
        command = shutil.which('catdoc')
        if command:
            output = subprocess.check_output(
                (command, '-s', '8859-1', path), cwd=os.path.dirname(command))
            print(sanitize_text(output.decode(encoding='iso-8859-1')))
        else:
            print('.doc preview requires catdoc.exe')
    elif ext in ('.doc', '.docx'):
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.visible = False
        word.Documents.Open(path)
        doc = word.ActiveDocument
        text = doc.Range().Text.replace("\x01", "").replace("\x07", "").replace("\r", "\n")
        text = re.sub("([ \t+]*\n){3,}", "\n", text)
        doc.Close()
        print(text.strip())
    else:
        print('Unknown file format')


def listdirs():
    """List configured directories"""
    for root in roots:
        print(root)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    default_config = os.path.join(
        appdirs.user_config_dir(), "finddoc", "finddoc.toml")
    parser.add_argument('--preview', action='store_const',
                        const=True, default=False)
    parser.add_argument('--config', default=default_config,
                        help=f'Path to config file (default: {default_config})')

    subs = parser.add_subparsers()
    subs.dest = 'command'
    subs.default = 'find'
    preview_parser = subs.add_parser('preview', help='Preview file')
    preview_parser.add_argument('file')
    update_parser = subs.add_parser('update', help='Update directory caches')
    find_parser = subs.add_parser('find', help='Find files (default)')
    list_parser = subs.add_parser('list', help='List included directories')

    opts = parser.parse_args()

    with open(opts.config, "rb") as infile:
        config = tomli.load(infile)
    roots = [parse_path(path) for path in config['finddoc']['paths']]

    if opts.command == 'list':
        listdirs()
    elif opts.command == 'find':
        fzf(opts)
    elif opts.command == 'preview':
        preview(opts.file)
    elif opts.command == 'update':
        update()
