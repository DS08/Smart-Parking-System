from setuptools import setup, find_packages
import py2exe

setup(
    console = ["Main.py"],
    version = "0.0.1",
    description = ("A simple module."),
    packages=find_packages(),
)
