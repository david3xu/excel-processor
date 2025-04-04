excel_processor/
├── __init__.py
├── main.py
├── cli.py
├── config.py
├── core/
│   ├── __init__.py
│   ├── reader.py
│   ├── structure.py
│   ├── extractor.py
├── models/
│   ├── __init__.py
│   ├── excel_structure.py
│   ├── metadata.py
│   ├── hierarchical_data.py
├── workflows/
│   ├── __init__.py
│   ├── base_workflow.py
│   ├── single_file.py
│   ├── multi_sheet.py
│   ├── batch.py
├── output/
│   ├── __init__.py
│   ├── formatter.py
│   ├── writer.py
├── utils/
│   ├── __init__.py
│   ├── caching.py
│   ├── exceptions.py
│   ├── logging.py
│   ├── progress.py
├── pyproject.toml
├── setup.py
├── README.md