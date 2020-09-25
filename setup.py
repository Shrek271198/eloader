from setuptools import setup, find_packages

setup(
    name='eload',
    version='0.1.0',
    license='proprietary',
    description='eMax Eload CLI',

    author='Devan Rehunathan',
    author_email='devan.rehunathan@technogen.com.au',
    url='https://www.technogen.com.au/',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=['pytest'],
    extras_require={},
    entry_points={
        'console_scripts': [
            'eloader = cli:main',
        ]
    },
)
