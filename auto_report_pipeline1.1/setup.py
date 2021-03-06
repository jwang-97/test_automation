from setuptools import setup

# 
setup(
    name="Auto_report",
    install_requires=[
        "python-pptx >= 0.6.19",
        "pypiwin32",
        "matplotlib >= 3.4.2",
        "pandas >= 1.3",
        "click >= 8.0",
    ],
    extras_require={
        "async": ["asgiref >= 3.2"],
        "dotenv": ["python-dotenv"],
    },
)
