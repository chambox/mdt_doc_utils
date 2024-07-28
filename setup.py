from setuptools import setup, find_packages

setup(
    name="mdt_doc_utils",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        "python-docx",
        "pytest"
    ],
)
