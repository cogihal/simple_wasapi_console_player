from comtypes import GUID


# Refer:
#   https://github.com/AndreMiras/pycaw/blob/develop/pycaw/constants.py
CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')


class EDataFlow:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/api/mmdeviceapi/ne-mmdeviceapi-edataflow
    eRender = 0
    eCompute = 1
    eAll = 2


class ERole:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/api/mmdeviceapi/ne-mmdeviceapi-erole
    eConsole = 0
    eMultimedia = 1
    eCommunications = 2


class DeviceState:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/coreaudio/device-state-xxx-constants
    ACTIVE = 0x01
    DISABLED = 0x02
    NOTPRESENT = 0x04
    UNPLUGGED = 0x08


class STGM:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/stg/stgm-constants
    STGM_READ = 0
    STGM_WRITE = 1
    STGM_READWRITE = 2


class AUDCLNT_SHAREMODE:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/api/audiosessiontypes/ne-audiosessiontypes-audclnt_sharemode
    AUDCLNT_SHAREMODE_SHARED = 0
    AUDCLNT_SHAREMODE_EXCLUSIVE = 1


# Refer:
#   https://learn.microsoft.com/ja-jp/windows/win32/coreaudio/audclnt-streamflags-xxx-constants
AUDCLNT_STREAMFLAGS_CROSSPROCESS        = 0x0001_0000
AUDCLNT_STREAMFLAGS_LOOPBACK            = 0x0002_0000
AUDCLNT_STREAMFLAGS_EVENTCALLBACK       = 0x0004_0000
AUDCLNT_STREAMFLAGS_NOPERSIST           = 0x0008_0000
AUDCLNT_STREAMFLAGS_RATEADJUST          = 0x0010_0000
AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM      = 0x8000_0000
AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY = 0x0800_0000

