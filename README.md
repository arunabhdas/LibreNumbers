# LibreNumbers

## Develop

```
source .venv/bin/activate

uv pip install -r requirements.txt
```

## Instantiating uv virtualenvironment

```
uv venv
```

## Pin python virtual environment with Python 3.11
```
uv venv --python 3.11
```

## Activate virtual environment

```
source .venv/bin/activate
```

## Install python-docx and lxml
```
uv pip install python-docx lxml
```

## Freeze dependencies

```
uv pip freeze > requirements.txt
```

## Recreate virtual environment

```
uv venv --python 3.11 .venv
````

## Usage

```
python libre_resume.py --in "/path/to/input.docx" --out "/path/to/output.docx"
```

## Usage

```
# 1) Normalize and show where the dangling '{' lives
python latex_to_docx_all_v2.py --in input.tex --emit-tex input.pandoc_ready.tex --debug-braces

# 2) Once brace imbalance is 0, run Pandoc (with LaTeX-like fonts via reference.docx)
python latex_to_docx_all_v2.py --in input.tex --out output.docx --run-pandoc --font-scheme cmu

# (Optional last resort) Automatically close missing braces, then run Pandoc
python latex_to_docx_all_v2.py --in input.tex --out output.docx --run-pandoc --font-scheme cmu --auto-fix-braces --debug-braces

```
