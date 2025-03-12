# Mass PPTX PDF Print

### Requirements

MS Windows, MS Powerpoint (Office)

### Why?

Open source pptx to pdf converters (pptx-python, pptxnextgen, node-pptx) sometimes cant convert stubborn files. It will be probably fine if you need to convert 5 files, but what if you have 500 files? Although python is slow as shit, this is easier than opening each file and running saveAs or print to pdf.

### Target Audience

Research projects with non-standard presentations, corporate workers with problematic keynotes and the like.

### Installation

```bash
poetry install
```

### Run

```bash
poetry run python <source-folder> [<target-folder>]
```
