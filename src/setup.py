from setuptools import find_packages, setup

setup(
    name="labware_tool",
    version="1.0.0-alpha-9",
    url="Github.com/TBD",
    license="BSD 3-clause",
    maintainer="Jet Mante",
    maintainer_email="jvm836@utexas.edu",
    include_package_data=True,
    description="creating ready to use labware defintions",
    packages=find_packages(include=["mantis_labware"]),
    long_description=open("README.md").read(),
    install_requires=[
        "pandas>=2.2.3",
        "openpyxl>=3.1.5",
    ],
    zip_safe=False,
)