# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

setup(
    name='NRCDataProxy',
    version='1.0',
    description="National Repsonse Center (NRC) Data Proxy",
    author='Tim Freund',
    author_email='tim@freunds.net',
    license = 'MIT',
    url='http://bitbucket.com/timfreund/NRCDataProxy',
    install_requires=[
        'flask',
        'pymongo'
                ],
    packages=['nrcdataproxy'],
    include_package_data=True,
    entry_points="""
    [console_scripts]
    nrcdataproxy = nrcdataproxy:serve
    """,
)
