from setuptools import setup

requirements = [
    "click==7.1.2",
    "openpyxl==3.0.5",
]

setup(
    name="eload",
    version="0.1.1",
    license="proprietary",
    description="eMax E-Load CLI",
    author="Devan Rehunathan",
    author_email="devan.rehunathan@technogen.com.au",
    url="https://www.technogen.com.au/",
    package_dir={'': 'src'},
    py_modules=["cli", "eload", "data", "writer", "utility"],
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "eloader = cli:main",
        ]
    },
)
