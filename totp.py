# -*- coding: utf-8 -*-

import hmac
import hashlib
import time
import base64
import functools


class Totp:
    step = 30

    def __init__(self, secret_key):
        try:
            if isinstance(secret_key, str):
                # Base32 decoding expects uppercase letters
                self.key = base64.b32decode(secret_key.upper())
            elif isinstance(secret_key, (bytes, bytearray)):
                self.key = secret_key
            else:
                raise ValueError("Secret key must be a string or bytes")
        except Exception as e:
            raise ValueError(f"Failed to decode secret key: {e}")
        
    def dynamic_truncate(self, digest_bytes: bytes) -> int:
        offset = digest_bytes[-1] & 0xf
        binary = int.from_bytes(digest_bytes[offset : offset + 4], byteorder='big')
        return binary & 0x7fffffff

    def generate_hotp(self, seed: bytes, counter: bytes, digits: int = 6, hash_algorithm=hashlib.sha1) -> str:
        digest_bytes = hmac.new(seed, counter, hash_algorithm).digest()
        otp = self.dynamic_truncate(digest_bytes) % (10 ** digits)
        return str(otp).zfill(digits)

    @staticmethod
    def get_current_unix_time() -> int:
        return int(time.time())

    @functools.lru_cache(maxsize=1)
    def get_current_steps(self) -> int:
        return self.get_current_unix_time() // self.step

    def generate_totp(self, digits: int = 6) -> str:
        steps_bytes = self.get_current_steps().to_bytes(8, byteorder='big')
        return self.generate_hotp(self.key, steps_bytes, digits)