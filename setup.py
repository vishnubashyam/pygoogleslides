from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pygoogleslides",
    version="0.1.0",
    author="Vishnu Bashyam",
    author_email="your.email@example.com",  # Replace with your email
    description="A Python package for automating Google Slides presentations",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/pygoogleslides",  # Replace with your repo URL
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ],
    python_requires=">=3.6",
    install_requires=[
        "google-auth>=2.0.0",
        "google-auth-oauthlib>=0.4.0",
        "google-auth-httplib2>=0.1.0",
        "google-api-python-client>=2.0.0",
    ],
) 