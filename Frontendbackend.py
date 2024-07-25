import os
import pytest

# Function to check if a file is an RTF file
def is_rtf_file(file_path):
    return os.path.splitext(file_path)[1].lower() == '.rtf'

# Function to check all files in a folder and print if they are RTF files
def check_folder_for_rtf(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            if is_rtf_file(file_path):
                print(f"{file_path} is an RTF file.")
            else:
                print(f"{file_path} is not an RTF file.")

# Test cases for the function
def test_is_rtf_file():
    # Test case 1: File with .rtf extension
    assert is_rtf_file("example.rtf") == True

    # Test case 2: File with .txt extension
    assert is_rtf_file("example.txt") == False

    # Test case 3: File with uppercase .RTF extension
    assert is_rtf_file("example.RTF") == True

    # Test case 4: File with no extension
    assert is_rtf_file("example") == False

# Main block to execute the function (for demonstration purposes)
if __name__ == "__main__":
    folder_path = "/Users/adithi/Desktop/jenkins_folder"  # Specify your folder path here
    check_folder_for_rtf(folder_path)
