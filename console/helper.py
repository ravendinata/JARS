from pathlib import Path

def get_source_file_path():
    """
    Prompts the user to enter the source file path and returns it.

    Returns:
        str: The source file path entered by the user.
    """

    file_path = input("\nEnter the source file path.\n(You don't need to remove the quotes (\"...\")): ")
    file_path = file_path.strip('\"')
    
    if not Path(file_path).exists():
        raise ValueError("Invalid file path.")
    
    return file_path

def get_output_file_path(source_file, extension):
    """
    Prompts the user to enter the output file path and returns it.

    Args:
        source_file (str): The path of the source file.

    Returns:
        str: The output file path entered by the user.
    """

    output_file_path = input("\nEnter the output file path (Leave blank to use same name and path as input): ")
    output_file_path = output_file_path.strip('\"')

    if not extension.startswith("."):
        extension = "." + extension

    if output_file_path == "":
        output_file_path = f"{source_file[:-4]}_processed{extension}"
    if not output_file_path.endswith(extension):
        output_file_path += extension
    
    return output_file_path

def prompt_overwrite(output_file_path):
    """
    Checks if the output file already exists and prompts the user for confirmation to overwrite it.

    Args:
        output_file_path (str): The path of the output file.

    Raises:
        ValueError: If the user chooses not to overwrite the file.
    """

    if Path(output_file_path).exists():
        overwrite = input("\nFile already exists. Do you want to overwrite the file? (y/n): ")
        if overwrite.lower() != "y":
            raise ValueError("File not overwritten.")