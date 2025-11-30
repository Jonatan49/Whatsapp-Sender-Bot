"""Setup script for WhatsApp Sender Bot Pro."""

from setuptools import setup, find_packages
import os

# Read README
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

# Read requirements
with open('requirements.txt', 'r', encoding='utf-8') as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

setup(
    name='whatsapp-sender-bot-pro',
    version='2.0.0',
    author='Yonatan Cohen',
    author_email='admin@whatsappbot.local',
    description='Professional WhatsApp bulk messaging bot with advanced features',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/yourusername/whatsapp-sender-bot-pro',
    packages=find_packages(),
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Topic :: Communications :: Chat',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
    ],
    python_requires='>=3.8',
    install_requires=requirements,
    entry_points={
        'console_scripts': [
            'whatsapp-bot=main:main',
        ],
    },
    include_package_data=True,
    package_data={
        '': ['config/*.yaml', 'assets/*'],
    },
)
