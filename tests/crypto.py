import os
from pathlib import Path

from cryptography.fernet import Fernet
from dotenv import load_dotenv


def get_key_from_env() -> bytes:
    """Retrieve the encryption key from the .env file.

    Raises:
        ValueError: If the environment variable ENCRYPTION_KEY is not found.

    Returns:
        The key as a bytes string.
    """
    load_dotenv()
    key = os.getenv("ENCRYPTION_KEY")
    if key is None:
        raise ValueError("Environment variable ENCRYPTION_KEY not found")
    return key.encode()  # The keys are binary, not strings.


def decrypt_file(encrypted_file: Path, decrypted_file: Path) -> None:
    key = get_key_from_env()
    fernet = Fernet(key)
    with open(encrypted_file, "rb") as file:
        encrypted_data = file.read()

    decrypted_data = fernet.decrypt(encrypted_data)
    with open(decrypted_file, "wb") as file:
        file.write(decrypted_data)


def encrypt_file(decrypted_file: Path, encrypted_file: Path) -> None:
    key = get_key_from_env()
    fernet = Fernet(key)
    with open(decrypted_file, "rb") as file:
        data = file.read()

    encrypted_data = fernet.encrypt(data)
    with open(encrypted_file, "wb") as file:
        file.write(encrypted_data)


def generate_key() -> None:
    key = Fernet.generate_key()
    with open(Path(".env"), "w") as env_file:
        env_file.write(f"ENCRYPTION_KEY={key.decode()}\n")
