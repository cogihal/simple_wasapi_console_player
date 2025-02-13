import time
import wave
from ctypes import POINTER, cast

import comtypes
from comtypes import CLSCTX_ALL, COMError
from pycaw.api.audioclient import WAVEFORMATEX
from pycaw.api.mmdeviceapi import PROPERTYKEY, IMMDeviceEnumerator
from pycaw.pycaw import IAudioClient

import core_audio_constants
from audio_render_client import IAudioRenderClient

import rlcompleter
from tab_complete import complete_path


def audio_device_id_list() -> tuple[list, list]:
    """
    Enumerate Core Audio devices and return a list of devices & GUIDs with the following process.

    1. CoInitialize()
    2. IMMDeviceEnumerator = CoCreateInstance(...)
    3. IMMDeviceCollection = IMMDeviceEnumerator::EnumAudioEndpoints(...)
    4. IMMDevice = IMMDeviceCollection::Item(i)
    5. id = IMMDevice::GetId()
    6. CoUninitialize()
    """

    comtypes.CoInitialize()

    device_enumerator = comtypes.CoCreateInstance(
        core_audio_constants.CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER,
    )

    collections = device_enumerator.EnumAudioEndpoints(
        core_audio_constants.EDataFlow.eRender,
        core_audio_constants.DeviceState.ACTIVE,
        # const.DeviceState.ACTIVE | const.DeviceState.UNPLUGGED,
    )

    devices = []
    ids = []

    count = collections.GetCount()
    for i in range(count):
        device = collections.Item(i)

        # Refer:
        #   https://github.com/AndreMiras/pycaw/blob/develop/pycaw/utils.py
        # 
        # property_store = device.OpenPropertyStore(const.STGM.STGM_READ)
        # property_count = property_store.GetCount()
        # for j in range(property_count):
        #     key = property_store.GetAt(j)
        #     val = property_store.GetValue(key)
        #     value = val.GetValue()
        #     print(f'{key=}, {val=}, {value=}')

        id = device.GetId()

        devices.append(device)
        ids.append(id)

    comtypes.CoUninitialize()

    return devices, ids


def default_audio_device_id() -> str:
    """
    Return the default audio device ID with the following process.

    1. CoInitialize()
    2. IMMDeviceEnumerator = CoCreateInstance(...)
    3. IMMDevice = IMMDeviceEnumerator::GetDefaultAudioEndpoint(...)
    4. id = IMMDevice::GetId()
    5. CoUninitialize()
    """

    comtypes.CoInitialize()

    device_enumerator = comtypes.CoCreateInstance(
        core_audio_constants.CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER,
    )

    audio_device = device_enumerator.GetDefaultAudioEndpoint(core_audio_constants.EDataFlow.eRender, core_audio_constants.ERole.eConsole)

    id = audio_device.GetId()

    comtypes.CoUninitialize()

    return id


def get_friendly_name(device_id) -> str:
    """
    Return the friendly name of the device from the device ID with the following process.

    1. CoInitialize()
    2. IMMDeviceEnumerator = CoCreateInstance(...)
    3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
    4. IPropertyStore = IMMDevice::OpenPropertyStore(STGM_READ)
    5. PROPERTYKEY = {A45C254E-DF1C-4EFD-8020-67D146A850E0}, 14
    6. value = IPropertyStore::GetValue(PROPERTYKEY)
    7. friendly_name = value.GetValue()
    8. CoUninitialize()
    """

    comtypes.CoInitialize()

    device_enumerator = comtypes.CoCreateInstance(
        core_audio_constants.CLSID_MMDeviceEnumerator,
        IMMDeviceEnumerator,
        comtypes.CLSCTX_INPROC_SERVER,
    )

    device = device_enumerator.GetDevice(device_id) # type: ignore
    property_store = device.OpenPropertyStore(core_audio_constants.STGM.STGM_READ)

    # Refer:
    #   https://github.com/AndreMiras/pycaw/blob/develop/pycaw/utils.py

    key = PROPERTYKEY()
    key.fmtid = comtypes.GUID('{A45C254E-DF1C-4EFD-8020-67D146A850E0}')
    key.pid = 14

    value = property_store.GetValue(comtypes.pointer(key))
    friendly_name = value.GetValue()

    comtypes.CoUninitialize()

    return friendly_name


from msvcrt import getch, kbhit


def get_key():
    if kbhit():
        ch = getch()
        # Space key
        if ch == b' ': return 1
        # q
        if ch == b'q': return -1
        return 0
    return 0


def main(audio_device, wavfile):
    wav = wave.open(wavfile, 'rb')

    frame_rate   = wav.getframerate() # frame rate (ex. 44100=44.1kHz)
    channels     = wav.getnchannels() # number of channels (monaural: 1, stereo: 2)
    sample_width = wav.getsampwidth() # sample byte size (ex. 16bit=2byte)
    frames       = wav.getnframes()   # number of frames
    data_size    = frames * channels * sample_width # data byte size

    comtypes.CoInitialize()

    interface = audio_device.Activate(
        IAudioClient._iid_, # type: ignore
        CLSCTX_ALL,
        None
    )
    audio_client = cast(interface, POINTER(IAudioClient))

    wav_format_ex = WAVEFORMATEX()
    wav_format_ex.wFormatTag      = 1
    wav_format_ex.nChannels       = channels
    wav_format_ex.nSamplesPerSec  = frame_rate
    wav_format_ex.wBitsPerSample  = sample_width * 8
    wav_format_ex.nBlockAlign     = wav_format_ex.nChannels * wav_format_ex.wBitsPerSample // 8
    wav_format_ex.nAvgBytesPerSec = wav_format_ex.nSamplesPerSec * wav_format_ex.nBlockAlign

    BUFFER_SIZE_IN_SECONDS = 2.0
    REFTIMES_PER_SEC = 10_000_000
    requestedSoundBufferDuration = int(REFTIMES_PER_SEC * BUFFER_SIZE_IN_SECONDS)
    stream_flags:int = core_audio_constants.AUDCLNT_STREAMFLAGS_RATEADJUST
    audio_client.Initialize(
        core_audio_constants.AUDCLNT_SHAREMODE.AUDCLNT_SHAREMODE_SHARED,
        stream_flags,
        requestedSoundBufferDuration,
        0,
        comtypes.pointer(wav_format_ex),
        None
    )

    service = audio_client.GetService(IAudioRenderClient._iid_) # type: ignore
    audio_render_client = cast(service, POINTER(IAudioRenderClient))

    num_buffer_frames = audio_client.GetBufferSize()

    audio_client.Start()

    play = True

    frame_chunk_size = 1024 # This number is changeable.
    wav_play_data = 0
    frames_to_write = 0

    while wav_play_data < data_size:
        try:
            # Check the keyboard operation.
            key = get_key()
            if key == -1:
                break
            if key == 1:
                play = not play

            if not play:
                time.sleep(0.1)
                continue

            # The padding frame means the data in the buffer and has not been played yet.
            buffer_padding_frames = audio_client.GetCurrentPadding()

            # If the buffer padding is less than frame_chunk_size, write frame_chunk_size frames.
            if buffer_padding_frames < frame_chunk_size:
                frames_to_write = frame_chunk_size
            else:
                # Enough data is in the buffer. Not need to write more. Wait for the buffer to be played.
                continue

            # Get the WASAPI buffer to play.
            buffer_to_play = audio_render_client.GetBuffer(frames_to_write)
            # Read the wav data.
            wav_frame_data = wav.readframes(frames_to_write)

            byte_len = len(wav_frame_data) # The byte size of actual read wav data.
            for i in range(byte_len):
                # Copy the wav data to the WASAPI buffer to play.
                buffer_to_play[i] = wav_frame_data[i]
            audio_render_client.ReleaseBuffer(frames_to_write, 0)
            wav_play_data += byte_len

        except COMError as e:
            # IAudioRenderClient can't be used anymore
            # possibly, the device is disconnected before finish playing
            break

    # print('Loop break') # _FOR_DEBUG_

    try:
        while audio_client.GetCurrentPadding() > 0:
            time.sleep(0.1)
    except COMError as e:
        pass

    # print('Finish playing') # _FOR_DEBUG_

    try:
        audio_client.Stop()
        audio_client.Release()
        audio_render_client.Release()
    except COMError as e:
        pass

    comtypes.CoUninitialize()


if __name__ == '__main__':
    defid = default_audio_device_id()
    devs, ids = audio_device_id_list()

    i = 0
    for id in ids:
        friendly_name = get_friendly_name(id)
        if id == defid:
            print(f'[*] {i} {friendly_name}')
        else:
            print(f'[ ] {i} {friendly_name}')
        i += 1

    devno = int(input('Select the device number: '))
    if devno < 0 or devno >= i:
        print('Invalid device number')
        exit(1)

    rlcompleter.readline.parse_and_bind("tab: complete")
    rlcompleter.readline.set_completer(complete_path)

    wavfile = input('Input the wave file name: ')
    import os
    if not os.path.exists(wavfile):
        print('Invalid file name')
        exit(1)

    print(f'q: quit, sp: pause/resume')

    main(devs[devno], wavfile)
