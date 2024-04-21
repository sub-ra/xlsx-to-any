from setuptools import setup, find_packages

setup(
    name='xlsx_to_any',
    version='1.0.0',
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'xlsx_to_any=xlsx_to_any:main',
        ],
    },
)
