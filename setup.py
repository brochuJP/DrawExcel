from setuptools import setup

def readme():
    with open('README.md') as f:
        README = f.read()
    return README

setup(
    name="DrawExcel",
    version="1.3",
    description="Tool that draws VBA structure of your MS Excel file.",
    long_description=readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/brochuJP/DrawExcel",
    author="Jean-Philippe Brochu",
    author_email="jpbrochu99@yahoo.ca",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
    ],
    packages=['DrawExcel'],
    include_package_data=True,
    install_requires=["pandas","win32com","graphviz"],
    #entry_points={
    #    "console_scripts": [
    #        "DrawExcel=DrawExcel:DrawExcel",
    #    ]
    #},
) 