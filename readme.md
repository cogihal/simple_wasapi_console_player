# Simple wav file console player using WASAPI

This is a simple wav file console player script using WASAPI.

## Features

- Select speaker to play.
- Select wav file to play.
- Pause, Resume, Quit playing by keyboard hit.


## Developing environments

- Windows 11 Pro
- Python 3.13.4  
- comtypes==1.4.11
- pycaw==20240210
- pyreadline3==3.5.4


## Caution

The following pycaw code should be edited.

pycaw/api/audioclient/depend.py

The type of the following variables should be changed from WORD to DWORD.
- nSamplesPerSec
- nAvgBytesPerSec

See also the following URL.

https://learn.microsoft.com/en-us/windows/win32/api/mmeapi/ns-mmeapi-waveformatex


## References

https://gist.github.com/kevinmoran/3d05e190fb4e7f27c1043a3ba321cede

