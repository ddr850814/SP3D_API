# -*- coding: utf-8 -*-
"""S3D 数据库登录文件的加解密。

算法逆向自 C:\\Win32app\\Smart3D\\ProjectMgmt\\Tools\\bin\\ConfigureDBLoginFile.exe
内置的 SP3DCoreDotNetUtils.S3dCryptServices：

  - 算法：RijndaelManaged（AES-256-CBC）+ PKCS7 填充
  - 默认口令（未调用 SetPassword / 传 null 时）："Juggernauts: awesome developers!"
  - key = SHA256(UTF8(口令)) 全部 32 字节
  - IV  = 同一 SHA256 摘要的后 16 字节（digest[16:32]）
  - 密文输出为 Base64

AES 内核用 cryptography 库实现（项目已依赖；原纯 Python 手写版经已知对
验证等价后退役）。已知验证对（曾在 Win 上用原始程序集复核）：
  Y6ZUoLAo4LsnyNJ5wEMklQ== -> pW4_s3admIn
  5IGkKN4JTsk98Vxb7Wh+6w== -> s3D807admin
"""
import base64
import hashlib

from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

DEFAULT_PASSWORD = "Juggernauts: awesome developers!"


def _derive_key_iv(password: str):
    digest = hashlib.sha256(password.encode("utf-8")).digest()
    return digest, digest[16:32]


def _pkcs7_unpad(data: bytes) -> bytes:
    # 与 S3D 原程序一致的宽容语义：非 PKCS7（如空填充）原样返回
    if not data:
        return data
    pad = data[-1]
    if pad < 1 or pad > 16 or data[-pad:] != bytes([pad]) * pad:
        return data
    return data[:-pad]


def s3d_encrypt(plain: str, password: str = DEFAULT_PASSWORD) -> str:
    """按 S3dCryptServices 算法加密，返回 Base64。"""
    key, iv = _derive_key_iv(password)
    raw = plain.encode("utf-8")
    pad = 16 - len(raw) % 16
    enc = Cipher(algorithms.AES(key), modes.CBC(iv)).encryptor()
    return base64.b64encode(enc.update(raw + bytes([pad]) * pad)
                            + enc.finalize()).decode("ascii")


def s3d_decrypt(b64: str, password: str = DEFAULT_PASSWORD):
    """按 S3dCryptServices 算法解密 Base64；失败返回 None。"""
    if not b64:
        return None
    key, iv = _derive_key_iv(password)
    try:
        data = base64.b64decode(b64)
    except Exception:
        return None
    if len(data) == 0 or len(data) % 16:
        return None
    dec = Cipher(algorithms.AES(key), modes.CBC(iv)).decryptor()
    out = dec.update(data) + dec.finalize()
    try:
        return _pkcs7_unpad(out).decode("utf-8")
    except Exception:
        return None


if __name__ == "__main__":
    pairs = [
        ("Y6ZUoLAo4LsnyNJ5wEMklQ==", "pW4_s3admIn"),
        ("5IGkKN4JTsk98Vxb7Wh+6w==", "s3D807admin"),
    ]
    for cipher, plain in pairs:
        got = s3d_decrypt(cipher)
        enc = s3d_encrypt(plain)
        print(f"dec({cipher}) = {got!r}  {'OK' if got == plain else 'FAIL'}")
        print(f"enc({plain}) = {enc}  {'OK' if enc == cipher else 'FAIL'}")
