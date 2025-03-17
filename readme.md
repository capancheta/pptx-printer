# Mass PPTX PDF Print

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)![Poetry](https://img.shields.io/badge/Poetry-%233B82F6.svg?style=for-the-badge&logo=poetry&logoColor=0B3D8D)![Windows](https://img.shields.io/badge/Windows-0078D6.svg?style=for-the-badge&logo=windows&logoColor=white)![Microsoft PowerPoint](https://img.shields.io/badge/Microsoft_PowerPoint-B7472A.svg?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white)

### Requirements

MS Windows, MS Powerpoint (Office)

### Why?

Open source pptx to pdf converters (pptx-python, pptxgenjs, nodejs-pptx, viewerjs) sometimes can't read/convert stubborn files. It will be probably fine if you need to convert 5 files, but what if you have 500 files? Although python is slow as shit, this beats opening each file and running saveAs or print to pdf.

### Target Audience

Research projects with dozens of non-standard presentations, corporate workers with repo-loads of problematic keynotes and the like.

### Installation

```bash
poetry install
```

### Run

Commandline:

```bash
poetry run python pptx-printer.py <source-folder> [<target-folder>]
```

gui mode:

```bash
poetry run python pptx-printer.py
```
