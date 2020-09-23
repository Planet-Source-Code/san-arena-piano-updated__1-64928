Attribute VB_Name = "Module1"
Option Explicit
Public Const MAXPNAMELEN = 32
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)
Public Const MIDIERR_BASE = 64
Public Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)
Public Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)
Public Const MIDIERR_BADOPENMODE = (MIDIERR_BASE + 6)
Type MIDIOUTCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
   wTechnology As Integer
   wVoices As Integer
   wNotes As Integer
   wChannelMask As Integer
   dwSupport As Long
End Type
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
