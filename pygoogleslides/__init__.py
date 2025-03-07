"""
PyGoogleSlides - A Python package for automating Google Slides presentations.
"""

from .auth import get_credentials, get_drive_service, get_slides_service
from .drive import (
    find_folder, 
    create_folder, 
    find_file, 
    delete_file, 
    rename_file, 
    copy_presentation, 
    move_file,
    find_or_create_folder
)
from .presentation import Presentation

__version__ = "0.1.0"
__author__ = "Vishnu Bashyam"

__all__ = [
    'get_credentials',
    'get_drive_service',
    'get_slides_service',
    'find_folder',
    'create_folder',
    'find_file',
    'delete_file',
    'rename_file',
    'copy_presentation',
    'move_file',
    'find_or_create_folder',
    'Presentation',
]