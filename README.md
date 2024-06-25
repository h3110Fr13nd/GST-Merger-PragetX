# GST Merger


## Installation and Setup

```bash
python -m venv .venv


source .venv/bin/activate
.venv/Scripts/activate


pip install -r requirements.txt
pyi-makespec --collect-data=gradio_client --collect-data=gradio --onefile main.py

```

Change `main.spec` to bypass pyz for gradio

```python
a = Analysis(
    ...
    module_collection_mode={
        'gradio': 'py',  # Collect gradio package as source .py files
    },
)
```

and make exe

```bash
pyinstaller main.spec
```

.exe file will be stored at dist folder
