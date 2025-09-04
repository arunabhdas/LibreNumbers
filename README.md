# LibreNumbers


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
