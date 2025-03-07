# PyGoogleSlides

A Python package for automating Google Slides presentations. This package provides a simple interface to create, modify, and manage Google Slides presentations programmatically.

## Installation

```bash
pip install pygoogleslides
```

## Prerequisites

1. A Google Cloud Project with the Google Slides API and Google Drive API enabled
2. A service account with appropriate permissions
3. The service account key file (JSON)

## Quick Start

```python
from pygoogleslides import (
    get_credentials,
    get_drive_service,
    get_slides_service,
    find_folder,
    create_folder,
    find_or_create_folder,
    find_file,
    delete_file,
    rename_file,
    copy_presentation,
    move_file,
    Presentation
)

# Configure your service account
SERVICE_ACCOUNT_FILE = 'path/to/your/service-account.json'
creds = get_credentials(SERVICE_ACCOUNT_FILE)
drive_service = get_drive_service(creds)
slides_service = get_slides_service(creds)

# Create a presentation from a template
shared_folder_id = find_folder(drive_service, 'Your Shared Folder Name')

# Create a hierarchical folder structure for organizing presentations
# This will find or create folders, and handle duplicates automatically 
grade_folder = find_or_create_folder(drive_service, 'Grade 7', shared_folder_id)
subject_folder = find_or_create_folder(drive_service, 'Math', grade_folder['id'])

# Find the template presentation
template_id = find_file(drive_service, 'Template Name', shared_folder_id)

# Create a new presentation in the subfolder, overwriting if it exists
new_presentation = copy_presentation(
    drive_service, 
    template_id, 
    'Unit 1 - Introduction to Algebra', 
    subject_folder['id'],
    overwrite=True  # This will delete any existing file with the same name
)

# Initialize the presentation object
presentation = Presentation(slides_service, new_presentation['id'])

# Replace text in the presentation
presentation.replace_text('{{placeholder}}', 'New Text')

# Replace an image
presentation.replace_image('{{image_placeholder}}', 'https://example.com/image.jpg')

# Add a hyperlink
presentation.replace_text('{{link_placeholder}}', 'Click Here', hyperlink='https://example.com')

# Rename the presentation if needed
renamed_presentation = rename_file(drive_service, new_presentation['id'], 'Renamed - Algebra Introduction')

# Create another folder and move the presentation if needed
another_subject = find_or_create_folder(drive_service, 'Science', grade_folder['id'])
moved_presentation = move_file(
    drive_service,
    new_presentation['id'],
    another_subject['id'],
    remove_parents=subject_folder['id']  # Remove from original folder
)
```

## Features

- Create copies of template presentations
- Create and manage folders for organizing presentations
- Handle duplicate folders by consolidating their contents
- Replace text placeholders
- Replace image placeholders
- Add hyperlinks
- Format text (bold, lists)
- Modify speaker notes
- Handle numbered lists and bullet points
- Rename presentations
- Move presentations between folders
- Delete and overwrite existing presentations

## API Reference

### Authentication

- `get_credentials(service_account_file, scopes=None)`: Get Google API credentials
- `get_drive_service(creds)`: Get Google Drive API service
- `get_slides_service(creds)`: Get Google Slides API service

### Drive Operations

- `find_folder(drive_service, folder_name, parent_folder_id=None, return_all=False)`: Find folders by name, optionally within a parent folder
- `create_folder(drive_service, folder_name, parent_folder_id=None)`: Create a new folder, optionally within a parent folder
- `find_or_create_folder(drive_service, folder_name, parent_folder_id=None)`: Find a folder or create it if it doesn't exist, handling duplicates by consolidating their contents
- `find_file(drive_service, file_name, parent_folder_id=None)`: Find a file by name, optionally within a parent folder
- `delete_file(drive_service, file_id)`: Delete a file
- `rename_file(drive_service, file_id, new_name)`: Rename a file
- `copy_presentation(drive_service, template_id, new_name, parents, overwrite=False)`: Create a copy of a presentation, optionally overwriting existing files with the same name
- `move_file(drive_service, file_id, new_parent_id, remove_parents=None)`: Move a file to a different folder

### Presentation Operations

The `Presentation` class provides methods for modifying presentations:

- `replace_text(placeholder, replacement, hyperlink=None, option_title=None)`: Replace text with optional hyperlink
- `replace_image(placeholder, image_url)`: Replace an image placeholder with a URL
- `create_slide(predefined_layout='BLANK')`: Create a new slide
- `delete_slide(slide_object_id)`: Delete a slide

## Examples

See the `examples` directory for complete usage examples:
- `folder_management_demo.py`: Demonstrates creating folders, moving files, and renaming presentations
- `folder_cleanup_demo.py`: Demonstrates how to handle duplicate folders and organize presentations in a hierarchical structure

## License

This project is licensed under the MIT License - see the LICENSE file for details.

