from setuptools import setup, find_packages

setup(
    name='OrcidXtract',  # Name of your package
    version='0.0.6',  # Starting version
    description='This Python library extracts ORCID information from a file containing ORCID IDs and generates reports in various formats (TXT, PDF, JSON, CSV, Excel).',
    long_description=open('src/README.md').read(),  # Read the long description from your README file
    long_description_content_type='text/markdown',
    author='Safial Islam Ayon',
    author_email='safialislam302@gmail.com',
    url='https://github.com/SafialIslam302/ORCID-Information',  # The URL to your project's repository
    packages=find_packages(where='src'),
    package_dir={'': 'src'},  # Tells setuptools that your code is in the `src/` folder
    install_requires=[  # Any dependencies your project requires
        'requests', 'pandas', 'reportlab~=4.3.1', 'openpyxl', 'pytest', 'coverage', 'typing_extensions', 'setuptools'
    ],
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',  # Or whatever license you're using
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',  # Minimum Python version required
    entry_points={
        'console_scripts': [
            'orcidxtract = OrcidXtract.main:main',
        ],
    },
)
