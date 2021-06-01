import setuptools

with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

setuptools.setup(
    name='xyw-macro',
    version='0.0.16',
    author='二炜',
    author_email='1174543101@qq.com',
    description='键盘宏，可根据需要自定义多套功能键区。',
    long_description=long_description,
    long_description_content_type='text/markdown',
    include_package_data=True,
    # package_data={
    #     'examples': ['*.csv', '*.xlsx', '*.exe'],
    #     'example/imgs': ['*.jpg']
    # },
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
    ],
    install_requires=[
        'pywin32',
        'simpleaudio'
    ],
)
