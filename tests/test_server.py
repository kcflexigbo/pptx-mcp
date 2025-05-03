import pytest
import os
from pathlib import Path

# Adjust the import path based on the project structure
# Assuming server.py is in the root directory and tests/ is a subdirectory
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent)) # Add project root to path

from server import (
    create_or_clear_presentation,
    add_slide,
    add_shape,
    get_slide_content_description,
    get_slide_image,
    SAVE_DIR,
    get_pptx_file,
)

# Define the test filename
TEST_FILENAME = "test_basic_create.pptx"
TEST_FILE_PATH = SAVE_DIR / TEST_FILENAME

def test_create_or_clear_presentation():
    """
    Tests if create_or_clear_presentation successfully creates a new file
    and returns the correct message.
    """
    # Ensure the file doesn't exist before the test
    if TEST_FILE_PATH.exists():
        os.remove(TEST_FILE_PATH)

    try:
        # --- Execute ---
        result = create_or_clear_presentation(TEST_FILENAME)

        # --- Verify ---
        # 1. Check if the file was created
        assert TEST_FILE_PATH.exists(), f"File '{TEST_FILE_PATH}' was not created."
        assert TEST_FILE_PATH.is_file(), f"'{TEST_FILE_PATH}' is not a file."

        # 2. Check the return message
        expected_message = f"Presentation '{TEST_FILENAME}' created/cleared successfully in '{SAVE_DIR}'."
        assert result == expected_message, f"Unexpected return message: {result}"

    finally:
        # --- Cleanup ---
        # Remove the test file after the test runs
        if TEST_FILE_PATH.exists():
            try:
                os.remove(TEST_FILE_PATH)
                # Optional: print statement for visibility during test runs
                # print(f"\nCleaned up test file: {TEST_FILE_PATH}")
            except OSError as e:
                print(f"Error removing test file {TEST_FILE_PATH}: {e}")

@pytest.mark.asyncio
async def test_add_slide_with_shapes():
    """
    Tests if we can create a slide with shapes and text.
    Verifies the shapes and their text content are added correctly.
    """
    # Ensure the file doesn't exist before the test
    if TEST_FILE_PATH.exists():
        os.remove(TEST_FILE_PATH)

    try:
        # --- Setup ---
        # Create a new presentation
        create_or_clear_presentation(TEST_FILENAME)
        
        # --- Execute ---
        # Add a blank slide (layout_index=6 is blank)
        add_slide(TEST_FILENAME, layout_index=6)
        
        # Add a rectangle with text
        add_shape(
            filename=TEST_FILENAME,
            slide_index=0,
            shape_type_name="RECTANGLE",
            left_inches=1.0,
            top_inches=1.0,
            width_inches=2.0,
            height_inches=1.0,
            text="Rectangle Text"
        )
        
        # Add an oval with text
        add_shape(
            filename=TEST_FILENAME,
            slide_index=0,
            shape_type_name="OVAL",
            left_inches=4.0,
            top_inches=1.0,
            width_inches=2.0,
            height_inches=1.0,
            text="Oval Text"
        )
        
        # --- Verify ---
        # Get the slide description to verify shapes and text
        description = await get_slide_content_description(TEST_FILENAME, "0")
        
        # Check if both shapes are present
        assert "Shape 0: Type=AUTO_SHAPE" in description, "Rectangle shape not found"
        assert "Shape 1: Type=AUTO_SHAPE" in description, "Oval shape not found"
        
        # Check if text content is correct
        assert "Text='Rectangle Text'" in description, "Rectangle text not found"
        assert "Text='Oval Text'" in description, "Oval text not found"
        
        # Check if positions are correct (within description)
        assert "Left=1.00\", Top=1.00\"" in description, "Shape position not correct"
        assert "Width=2.00\", Height=1.00\"" in description, "Shape dimensions not correct"

    finally:
        # --- Cleanup ---
        if TEST_FILE_PATH.exists():
            try:
                os.remove(TEST_FILE_PATH)
            except OSError as e:
                print(f"Error removing test file {TEST_FILE_PATH}: {e}")

@pytest.mark.asyncio
async def test_get_slide_image():
    """
    Tests if we can successfully get an image of a slide.
    Verifies that the image is returned and has valid PNG data.
    """
    # Ensure the file doesn't exist before the test
    if TEST_FILE_PATH.exists():
        os.remove(TEST_FILE_PATH)

    try:
        # --- Setup ---
        # Create a new presentation
        create_or_clear_presentation(TEST_FILENAME)
        
        # Add a blank slide (layout_index=6 is blank)
        add_slide(TEST_FILENAME, layout_index=6)
        
        # Add some shapes to make the slide visually interesting
        add_shape(
            filename=TEST_FILENAME,
            slide_index=0,
            shape_type_name="RECTANGLE",
            left_inches=1.0,
            top_inches=1.0,
            width_inches=2.0,
            height_inches=1.0,
            text="Test Rectangle"
        )
        
        add_shape(
            filename=TEST_FILENAME,
            slide_index=0,
            shape_type_name="OVAL",
            left_inches=4.0,
            top_inches=1.0,
            width_inches=2.0,
            height_inches=1.0,
            text="Test Oval"
        )
        
        # --- Execute ---
        # Get the image of the slide
        image = get_slide_image(TEST_FILENAME, 0)
        
        # --- Verify ---
        # Check that we got an Image object
        assert image is not None, "Image object should not be None"
        
        # Check that the image data is not empty
        assert len(image.data) > 0, "Image data should not be empty"
        
        # Check that the image format is PNG
        assert image._format == "png", "Image format should be PNG"
        
        # Check that the image data starts with PNG signature
        assert image.data.startswith(b'\x89PNG\r\n\x1a\n'), "Image data should be a valid PNG file"

    finally:
        # --- Cleanup ---
        if TEST_FILE_PATH.exists():
            try:
                os.remove(TEST_FILE_PATH)
            except OSError as e:
                print(f"Error removing test file {TEST_FILE_PATH}: {e}")

@pytest.mark.asyncio
async def test_get_pptx_file():
    """
    Tests if get_pptx_file returns a valid FileResource for a created presentation.
    """
    # Ensure the file doesn't exist before the test
    if TEST_FILE_PATH.exists():
        os.remove(TEST_FILE_PATH)

    try:
        # --- Setup ---
        create_or_clear_presentation(TEST_FILENAME)

        # --- Execute ---
        resource = await get_pptx_file(TEST_FILENAME)

        # --- Verify ---
        # Check that the resource is a FileResource
        from fastmcp.resources import FileResource
        assert isinstance(resource, FileResource), "Returned resource is not a FileResource"
        # Check that the file exists
        assert resource.path.exists(), f"File {resource.path} does not exist"
        # Check the mime type
        assert resource.mime_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        # Check the file signature (PPTX files are zip files, start with PK\x03\x04)
        file_bytes = resource.path.read_bytes()
        assert file_bytes[:4] == b'PK\x03\x04', "PPTX file does not start with correct ZIP signature"
    finally:
        # --- Cleanup ---
        if TEST_FILE_PATH.exists():
            try:
                os.remove(TEST_FILE_PATH)
            except OSError as e:
                print(f"Error removing test file {TEST_FILE_PATH}: {e}")
