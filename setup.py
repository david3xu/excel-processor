from setuptools import setup, find_packages

if __name__ == "__main__":
    setup(
        name="excel-processor",
        version="0.1.0",
        description="A comprehensive tool for processing Excel files with complex structures",
        author="Alcoa Team",
        packages=find_packages(),
        package_data={
            "excel_processor": ["config/*.json"],
        },
        include_package_data=True,
        entry_points={
            "console_scripts": [
                "excel-processor=cli:main",
            ],
        },
        install_requires=[
            "pandas",
            "openpyxl",
            "numpy",
            "pydantic",
        ],
        python_requires=">=3.8",
        classifiers=[
            "Development Status :: 4 - Beta",
            "Intended Audience :: Developers",
            "Programming Language :: Python :: 3",
            "Programming Language :: Python :: 3.8",
            "Programming Language :: Python :: 3.9",
            "Programming Language :: Python :: 3.10",
        ],
    )