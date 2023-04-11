# finddoc.py

`finddoc.py` is CLI tool that makes it easy and very fast to search and open
files on Windows when multiple project/document folders need to be searched.
Intended use is environments where the user may have personal Documents,
OneDrive, other folders, and network shares where documents (and other files)
are located. It handleds hundreds of thousands of file paths with ease.

Uses [fzf](https://github.com/junegunn/fzf) for the UI, and extends it with
support for multiple sources (not just one directory) and shortcuts for opening
a file with it's default association (<kbd>ENTER</kbd>), copying path to
clipboard (<kbd>alt-c</kbd>), and going to the file in explorer
(<kbd>alt-e</kbd>).

## Usage

### Interactively search for files

```
C:\> finddoc.py
```

### Shortcuts available when searching:

 * <kbd>ENTER</kbd> - open selected file (default associated program)
 * <kbd>alt-c</kbd> - copy full path to clipboard
 * <kbd>alt-e</kbd> - open Explorer with file selected
 * <kbd>alt-o</kbd> - open Total Commander with file selected
 * <kbd>alt-u</kbd> - update document cache (rescan all locations)
 * <kbd>ctrl-p/ctrl-n</kbd> - navigate history
 * additional standard fzf shortcuts work too


### Adding/removing directories

To add a path:

```bat
finddoc add <PATH>
```

For example, to add current path:

```bat
C:\MyProjects\Documents> finddoc add .
```

And to remote path:

```bat
finddoc remove <PATH>
```

`finddoc.py` will expand and normalize the argument, so relative paths will be handled correctly.


### Configure directories

The configuration file is in TOML format and must be placed in `%LOCALAPPDATA%\finddoc\finddoc.toml`.

Example (TOML format):

```toml
[finddoc]
paths = [
    "%USERPROFILE%\\Documents",
    "%ONEDRIVE%\\Documents",
    "P:\\Projects\\Fuschia\\Documents",
    "P:\\Projects\\WhiteGold\\Documents",
    "P:\\Guides\\Deployment",
    ...
]
```

### Update cache

To refresh directory caches:

```
C:\> finddoc.py update
```

### List directories

```
C:\> finddoc.py list
```
