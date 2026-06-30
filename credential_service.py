import ctypes
from ctypes import wintypes


CRED_TYPE_GENERIC = 1
CRED_PERSIST_LOCAL_MACHINE = 2


class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", wintypes.DWORD), ("dwHighDateTime", wintypes.DWORD)]


class CREDENTIALW(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD), ("Type", wintypes.DWORD), ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR), ("LastWritten", FILETIME), ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)), ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD), ("Attributes", ctypes.c_void_p),
        ("TargetAlias", wintypes.LPWSTR), ("UserName", wintypes.LPWSTR),
    ]


PCREDENTIALW = ctypes.POINTER(CREDENTIALW)


class CredentialService:
    PREFIX = "EmailAutomation:"

    def __init__(self):
        self.advapi = ctypes.WinDLL("Advapi32.dll")
        self.advapi.CredWriteW.argtypes = [PCREDENTIALW, wintypes.DWORD]
        self.advapi.CredWriteW.restype = wintypes.BOOL
        self.advapi.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(PCREDENTIALW)]
        self.advapi.CredReadW.restype = wintypes.BOOL
        self.advapi.CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
        self.advapi.CredDeleteW.restype = wintypes.BOOL
        self.advapi.CredFree.argtypes = [ctypes.c_void_p]

    def target(self, email):
        return self.PREFIX + email.strip().lower()

    def save_password(self, email, password):
        data = password.encode("utf-16-le")
        blob = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
        credential = CREDENTIALW()
        credential.Type = CRED_TYPE_GENERIC
        credential.TargetName = self.target(email)
        credential.CredentialBlobSize = len(data)
        credential.CredentialBlob = ctypes.cast(blob, ctypes.POINTER(ctypes.c_ubyte))
        credential.Persist = CRED_PERSIST_LOCAL_MACHINE
        credential.UserName = email
        if not self.advapi.CredWriteW(ctypes.byref(credential), 0):
            raise ctypes.WinError()

    def get_password(self, email):
        pointer = PCREDENTIALW()
        if not self.advapi.CredReadW(self.target(email), CRED_TYPE_GENERIC, 0, ctypes.byref(pointer)):
            return None
        try:
            credential = pointer.contents
            raw = ctypes.string_at(credential.CredentialBlob, credential.CredentialBlobSize)
            return raw.decode("utf-16-le")
        finally:
            self.advapi.CredFree(pointer)

    def delete_password(self, email):
        if not self.advapi.CredDeleteW(self.target(email), CRED_TYPE_GENERIC, 0):
            error = ctypes.get_last_error()
            if error not in (0, 1168):
                raise ctypes.WinError(error)
