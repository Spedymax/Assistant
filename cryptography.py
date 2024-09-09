import os
from fernet import Fernet


# Function to generate a key and save it into a file
def generate_key():
    key = Fernet.generate_key()
    with open("secret.key", "wb") as key_file:
        key_file.write(key)

# Function to load the key from a file
def load_key():
    return open("secret.key", "rb").read()

# Function to encrypt the file
def encrypt_file(file_name):
    # Generate a key or load an existing one
    if not os.path.exists("secret.key"):
        generate_key()
    key = load_key()

    # Initialize the Fernet class
    fernet = Fernet(key)

    # Read the Excel file
    with open(file_name, "rb") as file:
        original_file_data = file.read()

    # Encrypt the data
    encrypted_data = fernet.encrypt(original_file_data)

    # Write the encrypted data back to a new file
    with open("encrypted_" + file_name, "wb") as encrypted_file:
        encrypted_file.write(encrypted_data)

    print(f"{file_name} has been encrypted.")

# Function to decrypt the file
def decrypt_file(encrypted_file_name):
    key = load_key()

    # Initialize the Fernet class
    fernet = Fernet(key)

    # Read the encrypted file
    with open(encrypted_file_name, "rb") as encrypted_file:
        encrypted_data = encrypted_file.read()

    # Decrypt the data
    decrypted_data = fernet.decrypt(encrypted_data)

    # Write the decrypted data back to a new file
    decrypted_file_name = "decrypted_" + encrypted_file_name.replace("encrypted_", "")
    with open(decrypted_file_name, "wb") as decrypted_file:
        decrypted_file.write(decrypted_data)

    print(f"{encrypted_file_name} has been decrypted.")

# # Example usage
# file_name = "assistant_creds.xlsx"
# encrypt_file(file_name)

