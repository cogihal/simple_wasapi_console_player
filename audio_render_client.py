from ctypes import HRESULT, POINTER
from ctypes import c_ubyte as BYTE
from ctypes import c_uint32 as UINT32
from ctypes.wintypes import DWORD

from comtypes import COMMETHOD, GUID, IUnknown


class IAudioRenderClient(IUnknown):
    _iid_ = GUID("{F294ACFC-3146-4483-A7BF-ADDCA7C260E2}")
    _methods_ = (
        # HRESULT GetBuffer(
        # [in] UINT32 NumFramesRequested,
        # [out] BYTE **ppData);
        COMMETHOD(
            [],
            HRESULT,
            "GetBuffer",
            (["in"], UINT32, "NumFramesRequested"),
            (["out"], POINTER(POINTER(BYTE)), "ppData"),
        ),
        # HRESULT ReleaseBuffer(
        # [in] UINT32 NumFramesWritten,
        # [in] DWORD dwFlags);
        COMMETHOD(
            [],
            HRESULT,
            "ReleaseBuffer",
            (["in"], UINT32, "NumFramesWritten"),
            (["in"], DWORD, "dwFlags"),
        ),
    )

