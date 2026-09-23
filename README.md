# finddoc

Fast fuzzy document finder for Windows across multiple configured roots.

`finddoc` is a CLI tool for searching and opening files across personal folders, OneDrive locations, project directories, and network shares. It is designed for large path sets and uses [fzf](https://github.com/junegunn/fzf) for interactive selection.

## Features

- Search across multiple configured roots
- Open files with their default Windows association
- Copy selected file paths to the clipboard
- Reveal files in Explorer
- Optionally reveal files in Total Commander
- Cache scanned file lists for fast subsequent searches

## Installation

```powershell
uv sync
```

Ensure `fzf` is installed and available on `PATH`.

## Usage

### Interactively search for files

```powershell
uv run finddoc
```

### Search a specific directory

```powershell
uv run finddoc find C:\path\to\search
```

### Update the cache

```powershell
uv run finddoc update
```

### List configured directories

```powershell
uv run finddoc list
```

### Add a directory

```powershell
uv run finddoc add C:\path\to\include
```

### Remove a directory

```powershell
uv run finddoc remove C:\path\to\include
```

## Search shortcuts

- `Enter` — open selected file
- `Alt+C` — copy full path to clipboard
- `Alt+E` — open Explorer with file selected
- `Alt+O` — open Total Commander with file selected
- `Alt+U` — update document cache
- `Ctrl+P` / `Ctrl+N` — navigate history

## Configuration

The configuration file is stored at:

```text
%LOCALAPPDATA%\finddoc\finddoc.toml
```

Example:

```toml
[finddoc]
paths = [
    "%USERPROFILE%\\Documents",
    "%ONEDRIVE%\\Documents",
    "P:\\Projects\\Fuschia\\Documents",
    "P:\\Projects\\WhiteGold\\Documents",
]
```
