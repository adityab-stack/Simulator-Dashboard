"""
One-time script: generates an RSA key pair for Snowflake key-pair authentication.
Run this locally on your machine (not here) with: python gen_snowflake_key.py

Produces two files in the current folder:
  - snowflake_rsa_key.p8   (PRIVATE key - never commit this to git)
  - snowflake_rsa_key.pub  (PUBLIC key - safe to view, goes into a Snowflake ALTER USER command)
"""
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa

private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

private_bytes = private_key.private_bytes(
    encoding=serialization.Encoding.PEM,
    format=serialization.PrivateFormat.PKCS8,
    encryption_algorithm=serialization.NoEncryption(),
)
with open("snowflake_rsa_key.p8", "wb") as f:
    f.write(private_bytes)

public_bytes = private_key.public_key().public_bytes(
    encoding=serialization.Encoding.PEM,
    format=serialization.PublicFormat.SubjectPublicKeyInfo,
)
with open("snowflake_rsa_key.pub", "wb") as f:
    f.write(public_bytes)

print("Done. Created snowflake_rsa_key.p8 (PRIVATE - keep secret) and snowflake_rsa_key.pub (PUBLIC).")
print()
print("Public key body to paste into Snowflake's ALTER USER command (without BEGIN/END lines):")
print()
pub_text = public_bytes.decode()
body = "\n".join(pub_text.splitlines()[1:-1])
print(body)