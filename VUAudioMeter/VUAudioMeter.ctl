VERSION 5.00
Begin VB.UserControl VUAudioMeter 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   ToolboxBitmap   =   "VUAudioMeter.ctx":0000
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   2280
   End
   Begin VB.PictureBox picDEST 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   50
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   2
      Top             =   50
      Width           =   210
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   0
      Max             =   100
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSRC0 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   2520
      Picture         =   "VUAudioMeter.ctx":0312
      ScaleHeight     =   1935
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "VUAudioMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum xBorderStyle 'Controls Border Styles
    None = 0
    Inset = 1
    Raised = 2
    FixedSingle = 3
    Flat1 = 4
    Flat2 = 5
End Enum

Dim BSBorderStyle As xBorderStyle 'Controls BorderStyle

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020     'Copies the source bitmap to destination bitmap.
Private Const SRCAND = &H8800C6      'Combines pixels of the destination with source bitmap
                                    'using the Boolean AND operator.
Private Const SRCINVERT = &H660046   'Combines pixels of the destination with source bitmap
                                    'using the Boolean XOR operator.
Private Const SRCPAINT = &HEE0086    'Combines pixels of the destination with source bitmap
                                    'using the Boolean OR operator.
Private Const SRCERASE = &H4400328   'Inverts the destination bitmap and then combines the
                                    'results with the source bitmap using the Boolean AND
                                    'operator.
Private Const WHITENESS = &HFF0062   'Turns all output white.
Private Const BLACKNESS = &H42       'Turn output black.
 'This foreces all varibles to be declared now.
Dim i As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim hmixer As Long                      ' mixer handle
Dim inputVolCtrl As MIXERCONTROL        ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL       ' microphone volume control
Dim rc As Long                          ' return code
Dim OK As Boolean                       ' boolean return code
Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' Volume Buffer
Private VU As VULights                  ' Volume Unit Values

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim intLevel As Integer

Private VolValue As Double



Private Type VULights
    VUOn As Boolean
    InOutLev As Double
    VolLev As Double
    VULev As Double
    VolUnit As Variant
    VUArray As Long
    Freq(0 To 9) As Double
    FreqNum As Integer
    FreqVal As Double
End Type
Private Declare Sub ReleaseCapture Lib "user32" ()


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long

Private Const SND_SYNC = &H0 'just after the sound is ended exit private Function
Private Const SND_ASYNC = &H1 'just after the beginning of the sound exit private Function
Private Const SND_NODEFAULT = &H2 'if the sound cannot be found no error message
Private Const SND_LOOP = &H8 'repeat the sound until the private Function is called again
Private Const SND_NOSTOP = &H10 'if currently a sound is played the private Function will return without playing the selected sound
Private Const Flags& = SND_ASYNC Or SND_NODEFAULT
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Const CALLBACK_Function = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1         '  done bit
Private Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin

Private Type WAVEHDR
' The WAVEHDR user-defined type defines the header used to identify a waveform-audio buffer.
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   Reserved As Long        ' Not used
End Type

Private Type WAVEINCAPS
' The WAVEINCAPS user-defined variable describes the capabilities of a waveform-audio input
' device.
   wMid As Integer         ' Manufacturer identifier for the device driver for the
                           ' waveform-audio input device. Manufacturer identifiers
                           ' are defined in Manufacturer and Product Identifiers in
                           ' the Platform SDK product documentation.
   wPid As Integer         ' Product identifier for the waveform-audio input device.
                           ' Product identifiers are defined in Manufacturer and Product
                           ' Identifiers in the Platform SDK product documentation.
   vDriverVersion As Long  ' Version number of the device driver for the
                           ' waveform-audio input device. The high-order byte
                           ' is the major version number, and the low-order byte
                           ' is the minor version number.
   szPname As String * 32  ' Product name in a null-terminated string.
   dwFormats As Long       ' Standard formats that are supported. See the Platform
                           ' SDK product documentation for more information.
   wChannels As Integer    ' Number specifying whether the device supports
                           ' mono (1) or stereo (2) input.
End Type

Private Type WAVEFORMAT
' The WAVEFORMAT user-defined type describes the format of waveform-audio data. Only
' format information common to all waveform-audio data formats is included in this
' user-defined type.
   wFormatTag As Integer      ' Format type. Use the constant WAVE_FORMAT_PCM Waveform-audio data
                              ' to define the data as PCM.
   nChannels As Integer       ' Number of channels in the waveform-audio data. Mono data uses one
                              ' channel and stereo data uses two channels.
   nSamplesPerSec As Long     ' Sample rate, in samples per second.
   nAvgBytesPerSec As Long    ' Required average data transfer rate, in bytes per second. For
                              ' example, 16-bit stereo at 44.1 kHz has an average data rate of
                              ' 176,400 bytes per second (2 channels — 2 bytes per sample per
                              ' channel — 44,100 samples per second).
   nBlockAlign As Integer     ' Block alignment, in bytes. The block alignment is the minimum atomic unit of data. For PCM data, the block alignment is the number of bytes used by a single sample, including data for both channels if the data is stereo. For example, the block alignment for 16-bit stereo PCM is 4 bytes (2 channels — 2 bytes per sample).
   wBitsPerSample As Integer  ' For buffer estimation
   cbSize As Integer          ' Block size of the data.
End Type

Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
                                             ByVal uDeviceID As Long, _
                                             lpFormat As WAVEFORMAT, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal dwFlags As Long) As Long
' The waveInOpen private Function opens the given waveform-audio input device for recording. The private Function
' uses the following parameters
'     lphWaveIn-  a long value that is the handle identifying the open waveform-audio input
'                 device. Use this handle to identify the device when calling other
'                 waveform-audio input private Functions. This parameter can be NULL if WAVE_FORMAT_QUERY
'                 is specified for fdwOpen.
'     uDeviceID-  a long value that identifies the waveform-audio input device to open.
'                 This parameter can be either a device identifier or a handle of an open
'                 waveform-audio input device.
'     lpFormat-   the WAVEFORMAT user-defined typed that identifies the desired format for
'                 recording waveform-audio data.
'     dwCallback- a long value that is an event handle, a handle to a window, or the identifier
'                 of a thread to be called during waveform-audio recording to process messages
'                 related to the progress of recording. If no callback private Function is required,
'                 this value can be zero. For more information on the callback private Function,
'                 see waveInProc.
'     dwCallback- a long value that is the user-instance data passed to the callback mechanism.
'                 This parameter is not used with the window callback mechanism.
'     dwFlags-    Flags for opening the device. The following values are defined:
'                 CALLBACK_EVENT (&H50000)-event handle.
'                 CALLBACK_private Function (&H30000)-callback procedure address.
'                 CALLBACK_NULL (&H00000)-No callback mechanism. This is the default setting.
'                 CALLBACK_THREAD (&H20000)-thread identifier.
'                 CALLBACK_WINDOW (&H10000)-window handle.
'                 WAVE_FORMAT_DIRECT (&H8)-ACM driver does not perform conversions on the
'                                            audio data.
'                 WAVE_FORMAT_QUERY (&H1)-queries the device to determine whether it supports
'                                         the given format, but it does not open the device.
'                 WAVE_MAPPED (&H4)-The uDeviceID parameter specifies a waveform-audio device
'                                   to be mapped to by the wave mapper.

Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                      lpWaveInHdr As WAVEHDR, _
                                                      ByVal uSize As Long) As Long
' The waveInPrepareHeader private Function prepares a buffer for waveform-audio input. The private Function
' uses the following parameters:
'     hWaveIn-    a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-the WAVEHDR user-defined type variable.
'     uSize-      the size in bytes of the WAVEHDR user-defined type variable. Use the
'                 results of the Len private Function for this parameter.


Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInReset private Function stops input on the given waveform-audio input device and resets
' the current position to zero. All pending buffers are marked as done and returned to
' the application. The private Function requires the handle to the waveform-audio input device.

Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInStart private Function starts input on the given waveform-audio input device. The private Function
' requires the handle of the waveform-audio input device.

Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInStop private Function stops waveform-audio input. The private Function requires the handle of
' the waveform-audio input device.

Private Declare Function waveInUnprepareHeader Lib "winmm.dll" _
                                          (ByVal hWaveIn As Long, _
                                          lpWaveInHdr As WAVEHDR, _
                                          ByVal uSize As Long) As Long
' The waveInUnprepareHeader private Function cleans up the preparation performed by the
' waveInPrepareHeader private Function. This private Function must be called after the device driver
' fills a buffer and returns it to the application. You must call this private Function before
' freeing the buffer. The private Function uses the following parameters:
'     hWaveIn-       a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-   the variable typed as the WAVEHDR user-defined type identifying the
'                    buffer to be cleaned up.
'     uSize-         a long value that is the size in bytes, of the WAVEHDR varaible. Use
'                    the Len private Function with the WAVEHDR variable as the argument to get this
'                    value.

Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInClose private Function closes the given waveform-audio input device. The private Function
' requires the handle of the waveform-audio input device. If the private Function succeeds,
' the handle is no longer valid after this call.

Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                  (ByVal uDeviceID As Long, _
                  lpCaps As WAVEINCAPS, _
                  ByVal uSize As Long) As Long
' This private Function retrieves the capabilities of a given waveform-audio input device. You can
' use this private Function to determine the number of waveform-audio input devices present in the
' system. If the value specified by the uDeviceID parameter is a device identifier,
' it can vary from zero to one less than the number of devices present. The private Function uses
' the following parameters
'     uDeviceID-     long value that identifies waveform-audio output device. This value can be
'                    either a device identifier or a handle of an open waveform-audio input device.
'     lpCaps-user-   defined variable containing information about the capabilities of the device.
'     uSize-         the size in bytes of the user-defined variable used as the lpCaps parameter.
'                    Use the Len private Function to get this value.

Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
' The waveInGetNumDevs private Function returns the number of waveform-audio input devices present in the system.

Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
                     (ByVal err As Long, _
                     ByVal lpText As String, _
                     ByVal uSize As Long) As Long
'The waveInGetErrorText private Function retrieves a textual description of the error identified by
' the given error number. The private Function uses the following parameters:
'     Err-     a long value that is the error number.
'     lpText-  a string variable that contains the textual error description.
'     uSize-   the size in characters of the lpText string variable.

Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                   lpWaveInHdr As WAVEHDR, _
                                                   ByVal uSize As Long) As Long
' The waveInAddBuffer private Function sends an input buffer to the given waveform-audio input device.
' The private Function uses the following parameters:
'     hWaveIn-       a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-   the variable typed as the WAVEHDR user-defined type.
'     uSize-         a long value that is the size in bytes of the variable typed as the
'                    WAVEHDR user-defined variable. Use the Len private Function with the WAVEHDR
'                    variable as the argument to get this value.


Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32

Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&

Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)

Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)

Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
' The mixerClose private Function closes the specified mixer device. The private Function requires the
' handle of the mixer device. This handle must have been returned successfully by the
' mixerOpen private Function. If mixerClose is successful, the handle is no longer valid.

Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" _
            (ByVal hmxobj As Long, _
            pMxcd As MIXERCONTROLDETAILS, _
            ByVal fdwDetails As Long) As Long
' The mixerGetControlDetails private Function retrieves details about a single control associated
' with an audio line. the private Function uses the following parameters:
'     hmxobj-     a long value that is the handle to the mixer device object being queried.
'     pMxcd-      the variable defined as the MIXERCONTROLDETAILS user-defined type.
'     fdwDetails- Flags for retrieving control details. The following values are defined:
'                    MIXER_GETCONTROLDETAILSF_LISTTEXT-The paDetails member of the
'                       MIXERCONTROLDETAILS user-defined variable points to one or more
'                       MIXERCONTROLDETAILS_LISTTEXT user-defined variables to receive text
'                       Displays for multiple-item controls. An application must get all list
'                       text items for a multiple-item control at once. This flag cannot be
'                       used with MIXERCONTROL_CONTROLTYPE_CUSTOM controls.
'                    MIXER_GETCONTROLDETAILSF_VALUE-Current values for a control are
'                       retrieved. The paDetails member of the MIXERCONTROLDETAILS user-defined
'                       variable points to one or more details appropriate for the control class.
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs private Function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       The hmxobj parameter is the handle of a MIDI (Musical Instrument Digital
'                       Interface) input device. This handle must have been returned by the
'                       midiInOpen private Function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen private Function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The
'                       hmxobj parameter is a mixer device handle returned by the mixerOpen
'                       private Function. This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen private Function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen private Function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier
'                       of a MIDI input device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiInGetNumDevs
'                       private Function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier
'                       of a MIDI output device. This identifier must be in the range of zero
'                       to one less than the number of devices returned by the midiOutGetNumDevs
'                       private Function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of a
'                       mixer device in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs private Function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs private Function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs private Function.

Private Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" _
                  (ByVal uMxId As Long, _
                  ByVal pmxcaps As MIXERCAPS, _
                  ByVal cbmxcaps As Long) As Long
' The mixerGetDevCaps private Function queries a specified mixer device to determine its capabilities.
' The private Function uses the following parameters:
'     uMxId-      a long value that is the handle of an open mixer device.
'     pmxcaps-    a variable defined as the MIXERCAPS user-defined type to contain information
'                 about the capabilities of the device.
'     cbmxcaps-   a long value that is the size in bytes, of the variable defined as the
'                 MIXERCAPS user-defined type. Use the Len private Functions with the MIXERCAPS variable
'                 as the argument to get this value.

Private Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, _
                                             pumxID As Long, _
                                             ByVal fdwId As Long) As Long
' The mixerGetID private Function retrieves the device identifier for a mixer device associated
' with a specified device handle.The private Function uses the following parameters:
'     hmxobj-  a long value that is the handle of the audio mixer object to map to a
'              mixer device identifier.
'     pumxID-  the long value to contain the mixer device identifier. If no mixer device
'              is available for the hmxobj object, the value – 1 is placed in this location
'              and the MMSYSERR_NODRIVER error value is returned.
'     fdwId-   Flags for mapping the mixer object hmxobj. The following values are defined:
'                 MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs private Function.
'                 MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       the hmxobj parameter is the handle of a MIDI input device. This handle
'                       must have been returned by the midiInOpen private Function.
'                 MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen private Function.
'                 MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The hmxobj
'                       parameter is a mixer device handle returned by the mixerOpen private Function.
'                       This flag is optional.
'                 MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen private Function.
'                 MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen private Function.
'                 MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiInGetNumDevs
'                       private Function.
'                 MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiOutGetNumDevs
'                       private Function.
'                 MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of
'                       a mixer device in the range of zero to one less than the number of
'                       devices returned by the mixerGetNumDevs private Function. This flag is optional.
'                 MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs private Function.
'                 MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs private Function.

Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" _
                  (ByVal hmxobj As Long, _
                  pmxlc As MIXERLINECONTROLS, _
                  ByVal fdwControls As Long) As Long
' The mixerGetLineControls private Function retrieves one or more controls associated with an audio
' line. The private Function uses the following parameters:
'     hmxobj-        a long value that is the handle of the mixer device object that is being
'                    queried.
'     pmxlc-         the variable defined as the MIXERLINECONTROLS user-defined type used to
'                    reference one or more variables defined as theMIXERCONTROL user-defined
'                    types to be filled with information about the controls associated with
'                    an audio line. The cbStruct member of the MIXERLINECONTROLS variable
'                    must always be initialized to be the size, in bytes, of the
'                    MIXERLINECONTROLS variable.
'     fdwControls-   Flags for retrieving information about one or more controls associated w
'                    with an audio line. The following values are defined:
'                    MIXER_GETLINECONTROLSF_ALL-The pmxlc parameter references a list of
'                       MIXERCONTROL variables that will receive information on all controls
'                       associated with the audio line identified by the dwLineID member of
'                       the MIXERLINECONTROLS structure. The cControls member must be initialized
'                       to the number of controls associated with the line. This number is
'                       retrieved from the cControls member of the MIXERLINE structure returned
'                       by the mixerGetLineInfo private Function. The cbmxctrl member must be
'                       initialized to the size, in bytes, of a single MIXERCONTROL variable.
'                       The pamxctrl member must point to the first MIXERCONTROL variable to be
'                       filled. The dwControlID and dwControlType members are ignored for this
'                       query.
'                    MIXER_GETLINECONTROLSF_ONEBYID-The pmxlc parameter references a single
'                       MIXERCONTROL variable that will receive information on the control
'                       identified by the dwControlID member of the MIXERLINECONTROLS variable.
'                       The cControls member must be initialized to 1. The cbmxctrl member must
'                       be initialized to the size, in bytes, of a single MIXERCONTROL variable.
'                       The pamxctrl member must point to a MIXERCONTROL structure to be filled.
'                       The dwLineID and dwControlType members are ignored for this query. This
'                       query is usually used to refresh a control after receiving a
'                       MM_MIXM_CONTROL_CHANGE control change notification message by the
'                       user-defined callback (see mixerOpen).
'                    MIXER_GETLINECONTROLSF_ONEBYTYPE-The mixerGetLineControls private Function
'                       retrieves information about the first control of a specific class for
'                       the audio line that is being queried. The pmxlc parameter references a
'                       single MIXERCONTROL structure that will receive information about the
'                       specific control. The audio line is identified by the dwLineID member.
'                       The control class is specified in the dwControlType member of the
'                       MIXERLINECONTROLS variable. The dwControlID member is ignored for this
'                       query. This query can be used by an application to get information on
'                       a single control associated with a line. For example, you might want
'                       your application to use a peak meter only from a waveform-audio output
'                       line.
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs private Function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-The
'                       hmxobj parameter is the handle of a MIDI input device. This handle must
'                       have been returned by the midiInOpen private Function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen private Function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The
'                       hmxobj parameter is a mixer device handle returned by the mixerOpen
'                       private Function. This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen private Function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen private Function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiInGetNumDevs private Function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiOutGetNumDevs private Function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of a
'                       mixer device in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs private Function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs private Function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs private Function.

Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" _
                     (ByVal hmxobj As Long, _
                     pmxl As MIXERLINE, _
                     ByVal fdwInfo As Long) As Long
' The mixerGetLineInfo private Function retrieves information about a specific line of a mixer device.
' Uses the same parameters and constants as the mixerGetLineControls private Function.

Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
' The mixerGetNumDevs private Function retrieves the number of mixer devices present in the system.

Private Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, _
                                                ByVal uMsg As Long, _
                                                ByVal dwParam1 As Long, _
                                                ByVal dwParam2 As Long) As Long
' The mixerMessage private Function sends a custom mixer driver message directly to a mixer driver.
' The private Function uses the following parameters:
'     hmx-     a long value that is the handle of an open instance of a mixer device. This
'              value is the result of the mixerOpen private Function.
'     uMsg-    Custom mixer driver message to send to the mixer driver. This message must
'              be above or equal to the MXDM_USER constant.
'     dwParam1 and dwParam2-Arguments associated with the message being sent.

Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, _
                                             ByVal uMxId As Long, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal fdwOpen As Long) As Long
' The mixerOpen private Function opens a specified mixer device and ensures that the device will
' not be removed until the application closes the handle. the private Function uses the following
' parameters:
'     phmx-       a long value that is the handle identifying the opened mixer device. Use
'                 this handle to identify the device when calling other audio mixer private Functions.
'                 This parameter cannot be NULL.
'     uMxId-      a long value that identifies the mixer device to open. Use a valid device
'                 identifier or any HMIXEROBJ (see the mixerGetID private Function for a description of
'                 mixer object handles). A "mapper" for audio mixer devices does not currently
'                 exist, so a mixer device identifier of – 1 is not valid.
'     dwCallback- Handle of a window called when the state of an audio line and/or control
'                 associated with the device being opened is changed. Specify zero for this
'                 parameter if no callback mechanism is to be used.
'     dwInstance- User instance data passed to the callback private Function. This parameter is not
'                 used with window callback private Functions.
'     fdwOpen-    Flags for opening the device. The following values are defined:
'                    CALLBACK_WINDOW-  The dwCallback parameter is assumed to be a window handle.
'                    MIXER_OBJECTF_AUX (&H50000000)-The uMxId parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs private Function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-the
'                       uMxId parameter is the handle of a MIDI input device. This handle must
'                       have been returned by the midiInOpen private Function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       uMxId parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen private Function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The uMxId
'                       parameter is a mixer device handle returned by the mixerOpen private Function.
'                       This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       uMxId parameter is a waveform-audio input handle returned by the
'                       waveInOpen private Function.
'                    MIXER_OBJECTF_HWAVEOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       uMxId parameter is a waveform-audio output handle returned by the
'                       waveOutOpen private Function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000)-The uMxId parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiInGetNumDevs private Function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The uMxId parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiOutGetNumDevs private Function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The uMxId parameter is a mixer device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs private Function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The uMxId parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs private Function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The uMxId parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs private Function.

Private Declare Function mixerSetControlDetails Lib "winmm.dll" _
         (ByVal hmxobj As Long, _
         pMxcd As MIXERCONTROLDETAILS, _
         ByVal fdwDetails As Long) As Long
' The mixerSetControlDetails private Function sets properties of a single control associated with an
' audio line. The private Function uses the following parameters
'     hmxobj-        a long value that is the handle of the mixer device object for which
'                    properties are being set.
'     pMxcd-         the variable private Declares as the MIXERCONTROLDETAILS user-defined type.
'                    This variable references the control detail structures that contain the
'                    desired state for the control.
'     fdwDetails-    Flags for setting properties for a control. The following values are
'                    defined:
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs private Function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       The hmxobj parameter is the handle of a MIDI input device. This handle
'                       must have been returned by the midiInOpen private Function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen private Function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The hmxobj
'                       parameter is a mixer device handle returned by the mixerOpen private Function.
'                       This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen private Function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen private Function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000)-The hmxobj parameter is the identifier
'                       of a MIDI inputdevice. This identifier must be in the range of zero to
'                        one less than the number of devices returned by the midiInGetNumDevs
'                        private Function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier
'                       of a MIDI output device. This identifier must be in the range of zero
'                       to one less than the number of devices returned by the midiOutGetNumDevs
'                       private Function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is a mixer device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs private Function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs private Function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs private Function.
'                    MIXER_SETCONTROLDETAILSF_CUSTOM-A custom dialog box for the specified
'                       custom mixer control is displayed. The mixer device gathers the required
'                       information from the user and returns the data in the specified buffer.
'                       The handle for the owning window is specified in the hwndOwner member
'                       of the MIXERCONTROLDETAILS structure. (This handle can be set to NULL.)
'                       The application can then save the data from the dialog box and use it
'                       later to reset the control to the same state by using the
'                       MIXER_SETCONTROLDETAILSF_VALUE flag.
'                    MIXER_SETCONTROLDETAILSF_VALUE (&H00000000)-The current value(s) for a control
'                       are set. The paDetails member of the MIXERCONTROLDETAILS structure points
'                       to one or more mixer-control details structures of the appropriate class for
'                       the control.

Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
' The CopyStructFromPtr and CopyPtrFromStruct private Functions are user-defined versions of the
' RtlMoveMemory private Function. RtlMoveMemory moves memory either forward or backward, aligned or
' unaligned, in 4-byte blocks, followed by any remaining bytes. The private Function requires the
' following parameters:
'     Destination-   Pointer to the starting address of the copied block's destination.
'     Source-        Pointer to the starting address of the block of memory to copy.
'     Length-        Specifies the size, in bytes, of the block of memory to copy.

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                             ByVal dwBytes As Long) As Long
' The GlobalAlloc private Function allocates the specified number of bytes from the heap.
' Win32 memory management does not provide a separate local heap and global heap.
' This private Function is provided only for compatibility with 16-bit versions of Windows. The private Function
' uses the following parameters:
'     wFlags-     a long value that specifies how to allocate memory. If zero is specified,
'                 the default is allocate fixed memory. This value can be one or more of the
'                 following flags:
'                    GMEM_FIXED (&H0)- Allocates fixed memory. The return value is a pointer.
'                    GMEM_MOVEABLE (&H2)- Allocates movable memory. In Win32, memory blocks are
'                       never moved in physical memory, but they can be moved within the default .
'                       The return value is the handle of the memory object. To translate the
'                       heap handle into a pointer, use the GlobalLock private Function. This flag
'                       cannot be combined with the GMEM_FIXED flag.
'                    GPTR (GMEM_FIXED Or GMEM_ZEROINIT)-Combines the GMEM_FIXED and GMEM_ZEROINIT
'                       flags.
'                    GHND (GMEM_MOVEABLE Or GMEM_ZEROINIT)- Combines the GMEM_MOVEABLE and
'                       GMEM_ZEROINIT flags.
'                    GMEM_ZEROINIT (&H4)-Initializes memory contents to zero.
'     dwBytes-    Specifies the number of bytes to allocate. If this parameter is zero and
'                 the uFlags parameter specifies the GMEM_MOVEABLE flag, the private Function returns
'                 a handle to a memory object that is marked as discarded.

Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
' The GlobalLock private Function locks a global memory object and returns a pointer to the first
' byte of the object's memory block. This private Function is provided only for compatibility with
' 16-bit versions of Windows. The private Function requires a handle to the global memory object. This
' handle is returned by either the GlobalAlloc or GlobalReAlloc private Function.

Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
' The GlobalFree private Function frees the specified global memory object and invalidates its handle.
' This private Function is provided only for compatibility with 16-bit versions of Windows. The private Function
' requires a h andle to the global memory object. This handle is returned by either the
' GlobalAlloc or GlobalReAlloc private Function.

Private Type MIXERCAPS
' The MIXERCAPS user-defined type contains information about the capabilites of the mixer device.
   wMid As Integer                   '  manufacturer id
   wPid As Integer                   '  product id
   vDriverVersion As Long            '  version of the driver
   szPname As String * MAXPNAMELEN   '  product name
   fdwSupport As Long                '  misc. support bits
   cDestinations As Long             '  count of destinations
End Type

Private Type MIXERCONTROL
' The MIXERCONTROL user-defined type contains the state and metrics of a single control
' for an audio line.
   cbStruct As Long           '  size in Byte of MIXERCONTROL
   dwControlID As Long        '  unique control id for mixer device
   dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
   fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
   cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
   szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
   szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
   lMinimum As Long           '  Minimum value
   lMaximum As Long           '  Maximum value
   Reserved(10) As Long       '  reserved structure space
End Type

Private Type MIXERCONTROLDETAILS
' The MIXERCONTROLDETAILS user defined type refers to control-detail structures,
' retrieving or setting state information of an audio mixer control. All members of this
' user-defined type must be initialized before calling the mixerGetControlDetails and
' mixerSetControlDetails private Functions.
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

Private Type MIXERCONTROLDETAILS_SIGNED
' The MIXERCONTROLDETAILS_SIGNED user-defined type retrieves and sets signed type control
' properties for an audio mixer control.
   lValue As Long
End Type

Private Type MIXERLINE
' The MIXERLINE user-defined type describes the state and metrics of an audio line.
   cbStruct As Long        ' Size of MIXERLINE structure
   dwDestination As Long   ' Zero based destination index
   dwSource As Long        ' Zero based source index (if source)
   dwLineID As Long        ' Unique line id for mixer device
   fdwLine As Long         ' State/information about line
   dwUser As Long          ' Driver specific information
   dwComponentType As Long ' Component type for this audio line.
   cChannels As Long       ' Maximum number of separate channels that can be
                           ' manipulated independently for the audio line.
   cConnections As Long    ' Number of connections that are associated with the
                           ' audio line.
   cControls As Long       ' Number of controls associated with the audio line.
   szShortName As String * MIXER_SHORT_NAME_CHARS  ' Short string that describes
                                                   ' the audio mixer line specified
                                                   ' in the dwLineID member.
   szName As String * MIXER_LONG_NAME_CHARS  ' String that describes the audio
                                             ' mixer line specified in the dwLineID
                                             ' member. This description should be
                                             ' appropriate as a complete description
                                             ' for the line.
   dwType As Long          ' Target media device type associated with the audio
                           ' line described in the MIXERLINE structure.
   dwDeviceID As Long      ' Current device identifier of the target media device
                           ' when the dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
   wMid  As Integer        ' Manufacturer identifier of the target media device
                           ' when the dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
   wPid As Integer         ' Product identifier of the target media device when
                           ' the dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
   vDriverVersion As Long  ' Driver version of the target media device when the
                           ' dwType member is a target type other than
                           ' MIXERLINE_TARGETTYPE_UNDEFINED.
   szPname As String * MAXPNAMELEN  ' Product name of the target media device when
                                    ' the dwType member is a target type other than
                                    ' MIXERLINE_TARGETTYPE_UNDEFINED.
End Type

Private Type MIXERLINECONTROLS
' The MIXERLINECONTROLS user-defined type contains information about the controls
' of an audio line.
   cbStruct As Long     ' size in Byte of MIXERLINECONTROLS
   dwLineID As Long     ' Line identifier for which controls are being queried.
   dwControl As Long    ' Control identifier of the desired control
   cControls As Long    ' Number of MIXERCONTROL structure elements to retrieve.
   cbmxctrl As Long     ' Size, in bytes, of a single MIXERCONTROL structure.
   pamxctrl As Long     ' Address of one or more MIXERCONTROL structures to receive
                        '  the properties of the requested audio line controls.
End Type

'Private i As Integer
Private j As Integer
'private rc As Long
Private msg As String * 200
Private hWaveIn As Long
Private format As WAVEFORMAT

Private Const NUM_BUFFERS = 2
Private Const BUFFER_SIZE = 8192
Private Const DEVICEID = 0
Private hmem(NUM_BUFFERS) As Long
Private inHdr(NUM_BUFFERS) As WAVEHDR

Private fRecording As Boolean

Private Function GetControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
' This private Function attempts to obtain a mixer control. Returns True if successful.

   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   
   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   
   If (MMSYSERR_NOERROR = rc) Then
      mxlc.cbStruct = Len(mxlc)
      mxlc.dwLineID = mxl.dwLineID
      mxlc.dwControl = ctrlType
      mxlc.cControls = 1
      mxlc.cbmxctrl = Len(mxc)
      mxlc.pamxctrl = 9
      
      ' Allocate a buffer for the control
      'hmem = GlobalAlloc(&H40, Len(mxc))
      hmem = GlobalAlloc(GMEM_FIXED, Len(mxc))
      mxlc.pamxctrl = GlobalLock(hmem)
      mxc.cbStruct = Len(mxc)
      
      ' Get the control
      rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
            
      If (MMSYSERR_NOERROR = rc) Then
         GetControl = True
         
         ' Copy the control into the destination structure
         CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
      Else
         GetControl = False
      End If
      GlobalFree (hmem)
      Exit Function
   End If
   
   GetControl = False
End Function

' private Function to process the wave recording notifications.
Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

' This private Function starts recording from the soundcard. The soundcard must be recording in order to
' monitor the input level. Without starting the recording from this application, input level
' can still be monitored if another application is recording audio
Private Function StartInput() As Boolean

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

' Stop receiving audio input on the soundcard
Sub StopInput()

    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DrawBorder()
    'bit of credit to "Daniel Davies"
    Cls
    
    Select Case BSBorderStyle 'Draw The Border (If Any)
    
        Case 1 'Inset, We Need To Draw Several lines around the edge (8 to be exact)
            Line (0, 0)-(ScaleWidth, 0), vb3DDKShadow 'Darkest Shadow
            Line (1, 1)-(ScaleWidth - 1, 1), vb3DShadow 'Dark Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (1, 1)-(1, ScaleHeight - 1), vb3DShadow 'Dark Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2), vb3DLight 'Light Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 2, ScaleHeight - 2)-(1, ScaleHeight - 2), vb3DLight 'Light Shadow
            Refresh
        
        Case 2 'Raised, Same As Inset (But Colors Are Inverted)
            Line (0, 0)-(ScaleWidth, 0), vb3DHighlight 'Lightest Shadow
            Line (1, 1)-(ScaleWidth - 1, 1), vb3DLight 'Light Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (1, 1)-(1, ScaleHeight - 1), vb3DLight 'Light Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2), vb3DShadow 'Dark Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 2, ScaleHeight - 2)-(1, ScaleHeight - 2), vb3DShadow 'Dark Shadow
            Refresh
        
        Case 3 'Fixed Single (Black 1 Pixel Width Border)
            Line (0, 0)-(ScaleWidth, 0), vbBlack
            Line (0, 0)-(0, ScaleHeight), vbBlack
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vbBlack
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vbBlack
            Refresh
            
        Case 4 'Flat1 (Raised Then Inset)
            Line (0, 0)-(ScaleWidth, 0), vb3DHighlight 'Lightest Shadow
            Line (2, 2)-(ScaleWidth - 2, 2), vb3DDKShadow 'Darkest Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (2, 2)-(2, ScaleHeight - 2), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 3, 3)-(ScaleWidth - 3, ScaleHeight - 3), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 3, ScaleHeight - 3)-(1, ScaleHeight - 3), vb3DHighlight 'Lightest Shadow
            Refresh
        
        Case 5 'Flat2 (Inset Then Raised)
            Line (0, 0)-(ScaleWidth, 0), vb3DDKShadow 'Darkest Shadow
            Line (2, 2)-(ScaleWidth - 2, 2), vb3DHighlight 'Lightest Shadow
            Line (0, 0)-(0, ScaleHeight), vb3DDKShadow 'Darkest Shadow
            Line (2, 2)-(2, ScaleHeight - 2), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 3, 3)-(ScaleWidth - 3, ScaleHeight - 3), vb3DDKShadow 'Darkest Shadow
            Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DHighlight 'Lightest Shadow
            Line (ScaleWidth - 3, ScaleHeight - 3)-(1, ScaleHeight - 3), vb3DDKShadow 'Darkest Shadow
            Refresh
    
    End Select

End Function

Private Sub HScroll1_Change()

    SwapImages

End Sub

Private Sub HScroll1_Scroll()

    SwapImages

End Sub

Private Sub Timer1_Timer()
    VU.VolLev = volume / 327.67
    If (volume < 0) Then volume = -volume
    ' Get the current output level
    If (1 = 1) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    End If
    Value = VU.VolLev
End Sub

Private Sub UserControl_Initialize()
        Timer1.Interval = 50  '6.25
   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   ' Get the output volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
      ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
    Call BitBlt(picDEST.hDC, 0, 0, 13, 63, picSRC0.hDC, 1, 0, SRCAND)
'    picDEST.BorderStyle = 1

End Sub

Function SwapImages()

    Dim OneTwentyEighth As Long
    OneTwentyEighth = 100 / 28
    
    picDEST.Picture = Nothing
    i = HScroll1.Value / OneTwentyEighth
    
    Dim iY As Integer
        
    If i < 14 Then
iY = 1

    Else
        iY = 66
        i = i - 14
    End If
    
    Call BitBlt(picDEST.hDC, 0, 0, 13, 63, picSRC0.hDC, i * 15, iY, SRCAND)

    Label1.Caption = "" 'HScroll1.Value & "%"
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = Label1.ForeColor
'End Property
Public Property Get Enabled() As Boolean
    Enabled = Timer1.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Timer1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picSRC0,picSRC0,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = picSRC0.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set picSRC0.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=HScroll1,HScroll1,-1,Value
'Public Property Get Value() As Integer
'    Value = HScroll1.Value
'End Property

Public Property Let Value(ByVal New_Value As Integer)
Attribute Value.VB_Description = "Returns/sets the value of an object."
    HScroll1.Value() = New_Value
    PropertyChanged "Value"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BorderStyle = PropBag.ReadProperty("BorderStyle", xBorderStyle.Flat1)
    Enabled = PropBag.ReadProperty("Enabled", False)
'    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &HFFC0FF)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
'    HScroll1.Value = PropBag.ReadProperty("Value", 0)
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .Height = 1045
        .Width = 305
    End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", BSBorderStyle, xBorderStyle.Flat1)
    Call PropBag.WriteProperty("Enabled", Timer1.Enabled, False)
'    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &HFFC0FF)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
'    Call PropBag.WriteProperty("Value", HScroll1.Value, 0)
End Sub

Public Property Get BorderStyle() As xBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = BSBorderStyle  'Change The Value
End Property

Public Property Let BorderStyle(ByVal NewStyle As xBorderStyle)
    BSBorderStyle = NewStyle 'Change The BorderStyle
    DrawBorder 'Redraw The Border
    SwapImages
End Property
