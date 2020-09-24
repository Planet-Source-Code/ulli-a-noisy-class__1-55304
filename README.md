<div align="center">

## A Noisy Class


</div>

### Description

Simple Sound Mixer Wrapper Class plus Test Driver. Unfortunately the Mixer interface is rather complicated - maybe written by a musician *g* - so there are quite a few mystic API calls with plenty of params, mixer-constants with ugly names, cryptic structure types, and virtual memory address pointers from one structure into the next. And Micro$oft's documentation is slim, to put it polite, but I tried my best to put all that into a wrapper with a simple interface: You choose the Channel and SoundControl; this will return True if the selection was successful, and then Get or Let the Value. (Note that ALL values are in % - for booleans (like Mute) the value 0 means False and 100 means True - one hundred percent true, so to say).
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2004-08-02 10:00:10
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[A\_Noisy\_Cl177678822004\.zip](https://github.com/Planet-Source-Code/ulli-a-noisy-class__1-55304/archive/master.zip)

### API Declarations

```
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Enum MixerConstants 'this makes them long by default
  MMSYSERR_NOERROR = 0
  MIXER_CONTROLDETAILSF_VALUE = 0
  MIXER_GETLINECONTROLSF_ONEBYTYPE = 2
  MIXER_GETLINEINFOF_COMPONENTTYPE = 3
  MIXER_SHORT_NAME_CHARS = 16
  MIXER_LONG_NAME_CHARS = 64
  MAXPNAMELEN = 32
  'mixer line constants
  MLC_DST_FIRST = 0
  MLC_SRC_FIRST = &H1000
  'Mixer control constants
  MCT_CLASS_FADER = &H50000000
  MCT_UNITS_UNSIGNED = &H30000
  MCT_FADER = MCT_CLASS_FADER Or MCT_UNITS_UNSIGNED
  MCT_CLASS_SWITCH = &H20000000
  MCT_UNITS_BOOLEAN = &H10000
  MCT_BOOLEAN = MCT_CLASS_SWITCH Or MCT_UNITS_BOOLEAN
End Enum
Public Enum Channels
  DigitalOut = MLC_DST_FIRST + 1
  LineOut = MLC_DST_FIRST + 2
  MonitorOut = MLC_DST_FIRST + 3
  SpeakersOut = MLC_DST_FIRST + 4
  HeadphonesOut = MLC_DST_FIRST + 5
  TelephoneOut = MLC_DST_FIRST + 6
  WaveInOut = MLC_DST_FIRST + 7
  VoiceInOut = MLC_DST_FIRST + 8
  DigitalIn = MLC_SRC_FIRST + 1
  LineIn = MLC_SRC_FIRST + 2
  MikrophoneIn = MLC_SRC_FIRST + 3
  SynthesizerIn = MLC_SRC_FIRST + 4
  CompactDiscIn = MLC_SRC_FIRST + 5
  TelephoneIn = MLC_SRC_FIRST + 6
  PCSpeakerIn = MLC_SRC_FIRST + 7
  WaveOutIn = MLC_SRC_FIRST + 8
  AuxiliaryIn = MLC_SRC_FIRST + 9
  AnalogIn = MLC_SRC_FIRST + 10
End Enum
#If False Then
Private DigitalOut, LineOut, MonitorOut, SpeakersOut, HeadphonesOut, TelephoneOut, WaveInOut, VoiceInOut
Private DigitalIn, LineIn, MikrophoneIn, SynthesizerIn, CompactDiscIn, TelephoneIn, PCSpeakerIn, WaveOutIn, AuxiliaryIn, AnalogIn
#End If
Public Enum SoundControls
  Loudness = MCT_BOOLEAN + 4
  Mute = MCT_BOOLEAN + 2
  StereoEnhance = MCT_BOOLEAN + 5
  Mono = MCT_BOOLEAN + 3
  Volume = MCT_FADER + 1
  Bass = MCT_FADER + 2
  Treble = MCT_FADER + 3
  Equalizer = MCT_FADER + 4
End Enum
#If False Then
Private Loudness, Mute, StereoEnhance, Mono, Pan, Volume, Bass, Treble, Equalizer
#End If
'mixer handle
Private hMixer As Long
'mixer structures
Private Type MIXERLINE
  cbStruct      As Long 'size in bytes of MIXERLINE structure
  dwDestination    As Long 'zero based destination index
  dwSource      As Long 'zero based source index (if source)
  dwLineID      As Long 'unique line id for mixer device
  fdwLine       As Long 'state/information about line
  dwUser       As Long 'driver specific information
  dwComponentType   As Long 'component type line connects to
  cChannels      As Long 'number of channels line supports
  cConnections    As Long 'number of connections (possible)
  cControls      As Long 'number of controls at this line
  szShortName(1 To MIXER_SHORT_NAME_CHARS)  As Byte
  szName(1 To MIXER_LONG_NAME_CHARS)     As Byte
  dwType       As Long
  dwDeviceID     As Long
  wMid        As Integer
  wPid        As Integer
  vDriverVersion   As Long
  szPname(1 To MAXPNAMELEN) As Byte
End Type
Private ChannelLine As MIXERLINE
Private Type MIXERLINECONTROLS
  cbStruct      As Long 'size in Byte of MIXERLINECONTROLS
  dwLineID      As Long 'line id (from MIXERLINE.dwLineID)
  dwControl      As Long 'MIXER_GETLINECONTROLSF_ONEBYID or MIXER_GETLINECONTROLSF_ONEBYTYPE
  cControls      As Long 'count of controls pamxctrl points to
  cbmxctrl      As Long 'size in Byte of _one_ MIXERCONTROL
  pamxctrl      As Long 'pointer to first MIXERCONTROL array
End Type
Private ChannelControls As MIXERLINECONTROLS
Private Type MIXERCONTROL
  cbStruct      As Long 'size in Byte of MIXERCONTROL
  dwControlID     As Long 'unique control id for mixer device
  dwControlType    As Long 'MIXERCONTROL_CONTROLTYPE_xxx
  fdwControl     As Long 'MIXERCONTROL_CONTROLF_xxx
  cMultipleItems   As Long 'if MIXERCONTROL_CONTROLF_MULTIPLE set
  szShortName(1 To MIXER_SHORT_NAME_CHARS)  As Byte 'short name of control
  szName(1 To MIXER_LONG_NAME_CHARS)     As Byte 'long name of control
  lMinimum      As Long 'Minimum value
  lMaximum      As Long 'Maximum value
  reserved(10)    As Long 'reserved structure space
End Type
Private ValueControl As MIXERCONTROL
Private Type MIXERCONTROLDETAILS
  cbStruct      As Long 'size in Byte of MIXERCONTROLDETAILS
  dwControlID     As Long 'control id to get/set details on
  cChannels      As Long 'number of channels in paDetails array
  item        As Long 'hwndOwner or cMultipleItems
  cbDetails      As Long 'size of one details_XX struct
  paDetails      As Long 'pointer to array of details_XX structs
End Type
Private ControlDetails As MIXERCONTROLDETAILS
```





