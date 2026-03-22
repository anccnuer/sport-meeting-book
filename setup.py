from setuptools import setup, find_packages

setup(
    name='sports-meet-order-book',
    version='1.0.0',
    description='运动会秩序册生成器',
    author='Your Name',
    author_email='your.email@example.com',
    packages=find_packages('src'),
    package_dir={'': 'src'},
    install_requires=[
        'pandas==2.3.3',
        'openpyxl==3.1.5',
        'python-docx==1.2.0',
        'lxml==5.3.1'
    ],
    entry_points={
        'console_scripts': [
            'sports-meet=main:main'
        ]
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.8',
)