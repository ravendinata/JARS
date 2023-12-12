import datetime
import hashlib

from termcolor import colored

@staticmethod
def hash_file(path):
    """
    Hashes a file using blake2b algorithm.
    """
    # Make a hash object
    hash = hashlib.blake2b()
    
    # Open file for reading in binary mode
    with open(path, "rb") as file:
        # Loop till the end of the file
        chunk = 0
        while chunk != b'':
            # Read only 1024 bytes at a time
            chunk = file.read(1024)
            hash.update(chunk)
    
    # return the hex representation of digest
    return hash.hexdigest()

@staticmethod
def generate_serial_number(file_path):
    """
    Generates a serial number based on the current date, 
    first 8 characters of the hash of the file, 
    and last 8 characters of the hash of the file.
    """
    # Get the current date
    date = datetime.datetime.now().strftime("%Y%m%d")
    
    # Get the hash of the file
    hash = hash_file(file_path)
    
    # Return the serial number
    return f"{date}-{hash[:8]}-{hash[-8:]}"

@staticmethod
def sign_pdf(file_path, verbose = False):
    """
    Signs a PDF file using the JAC Digital Signature.
    """
    
    from pypdf import PdfWriter, PdfReader

    if verbose:
        print(f"Signing file: {file_path}")

    reader = PdfReader(file_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    metadata = reader.metadata
    writer.add_metadata(metadata)
    writer.add_metadata(
        {
            "/Serial Number": generate_serial_number(file_path)
        }
    )

    with open(file_path, "wb") as file:
        writer.write(file)

    hash = hash_file(file_path)

    reader = PdfReader(file_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    metadata = reader.metadata
    writer.add_metadata(metadata)
    writer.add_metadata(
        {
            "/Signed By": "JARS InManage",
            "/Signed On": datetime.datetime.now().strftime(f"%Y/%m/%d @ %H:%M:%S"),
            "/Signature": "JARS InManage Digital Signature (JARSIM)",
            "/Hash": hash
        }
    )

    if verbose:
        print(f"Adding metadata:\n{metadata}")

    with open(file_path, "wb") as file:
        writer.write(file)

    return hash

@staticmethod
def verify_pdf(file_path, verbose = False):
    """
    Verifies the integrity of a PDF file.
    """
    import os
    from pypdf import PdfReader, PdfWriter

    print(colored(f"\nv^ Verifying file: {file_path}", "white", "on_magenta"))

    # 1. Check if the hash matches
    print(colored("v^ Step 1 > Checking Hash", "magenta"))
    reader = PdfReader(file_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    metadata = reader.metadata

    try:
        stored_hash = metadata.pop("/Hash")
        metadata.pop("/Signed By")
        metadata.pop("/Signed On")
        metadata.pop("/Signature")
    except:
        output_text = "The file is not signed. It is likely that the file is not a JARS report or it may have been severely tampered with."
        print(colored("(!)", "red"),
              colored(output_text, "white", "on_red"))
        return False, output_text

    writer.add_metadata(metadata)

    temp_file = os.path.splitext(file_path)[0]
    writer.write(f"{temp_file}.temp.pdf")

    computed_hash = hash_file(f"{temp_file}.temp.pdf")
    
    if verbose:
        print(f"Stored hash: {stored_hash[:8]} vs Computed hash: {computed_hash[:8]}")
    
    if stored_hash != computed_hash:
        output_text = "The file is signed but the hash does not match. It is very likely that it has been tampered with."
        print(colored("(!)", "red"),
              colored(output_text, "white", "on_red"))
        os.remove(f"{temp_file}.temp.pdf")
        return False, output_text
    else:
        print(colored("v^", "green"), 
              colored("The file is signed and the hash matches. This file is integrous and is highly unlikely to have been tampered with.", "white", "on_green"))
    
    # 2. Check if the serial number matches
    print(colored("v^ Step 2 > Checking Serial number", "magenta"))
    tf_reader = PdfReader(f"{temp_file}.temp.pdf")
    tf_writer = PdfWriter()

    for page in tf_reader.pages:
        tf_writer.add_page(page)
    
    tf_metadata = tf_reader.metadata
    try:
        sn = tf_metadata.pop("/Serial Number")
    except:
        output_text = "The file is signed but the serial number is missing. It is very likely that it has been tampered with."
        print(colored("(!)", "red"),
              colored(output_text, "white", "on_red"))
        return False, output_text

    tf_writer.add_metadata(tf_metadata)
    tf_writer.write(f"{temp_file}.temp.pdf")

    sn_split = sn.split("-")
    sn_hash_start = sn_split[1]
    sn_hash_end = sn_split[2]
    sn_computed_hash_start = hash_file(f"{temp_file}.temp.pdf")[:8]
    sn_computed_hash_end = hash_file(f"{temp_file}.temp.pdf")[-8:]

    os.remove(f"{temp_file}.temp.pdf")

    if verbose:
        print(f"Stored serial number: {sn_hash_start}...{sn_hash_end} vs Computed serial number: {sn_computed_hash_start}...{sn_computed_hash_end}")

    if sn_hash_start != sn_computed_hash_start or sn_hash_end != sn_computed_hash_end:
        output_text = "The file is signed but the serial number does not match. It is very likely that it has been tampered with."
        print(colored("(!)", "red"),
              colored(output_text, "white", "on_red"))
        return False, output_text
    else:
        output_text = "The file is signed, the hash matches, and the serial number matches. This file is integrous and is highly unlikely to have been tampered with."
        print(colored("v^", "green"), 
              colored(output_text, "white", "on_green"))
        return True, output_text