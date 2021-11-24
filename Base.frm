VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Base 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deeth Spectrum Analyzer"
   ClientHeight    =   1770
   ClientLeft      =   1440
   ClientTop       =   2490
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider Slider 
      Height          =   228
      Left            =   2196
      TabIndex        =   4
      Top             =   1440
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   397
      _Version        =   393216
      Enabled         =   0   'False
      LargeChange     =   1
      Min             =   1
      SelStart        =   5
      Value           =   5
   End
   Begin VB.Timer QuitTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2592
      Top             =   1368
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000002&
      Height          =   816
      Left            =   72
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   2
      Top             =   468
      Width           =   3120
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   336
      Left            =   1152
      TabIndex        =   3
      Top             =   1368
      Width           =   984
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   288
      Left            =   72
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   72
      Width           =   3108
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "&Start"
      Height          =   336
      Left            =   72
      TabIndex        =   1
      Top             =   1368
      Width           =   984
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   336
      Left            =   2196
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   5
      Top             =   1368
      Visible         =   0   'False
      Width           =   336
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private DevHandle As Long 'Handle of the open audio device

Private Visualizing As Boolean
Private Divisor As Long

Private ScopeHeight As Long 'Saves time because hitting up a Long is faster
                            'than a property.
                            
Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long


Sub InitDevices()
    'Fill the DevicesBox box with all the compatible audio input devices
    'Bail if there are none.
    
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        If Caps.Formats And WAVE_FORMAT_1M16 Then '16-bit mono devices
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End 'Ewww!  End!  Bad me!
    End If
    DevicesBox.ListIndex = 0
End Sub


Private Sub Form_Load()
    Call InitDevices 'Fill the DevicesBox
    
    Call DoReverse   'Pre-calculate these
    
    Call Slider_Change 'Initialize this
    
    'Set the double buffer to match the display
    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    
    ScopeHeight = Scope.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If DevHandle <> 0 Then
        Call DoStop
        Cancel = 1
        If Visualizing = True Then
            QuitTimer.Enabled = True
        End If
    End If
End Sub


Private Sub QuitTimer_Timer()
    Unload Me
End Sub


Private Sub Slider_Change()
    'This essentually adjusts the scale of the spectrum
    Divisor = ((Slider.Max - Slider.Value + 1) / Slider.Max) * 5200
End Sub


Private Sub Slider_Scroll()
    Call Slider_Change
End Sub

Private Sub StartButton_Click()
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    StopButton.Enabled = True
    StartButton.Enabled = False
    Slider.Enabled = True
    DevicesBox.Enabled = False
    
    Call Visualize
End Sub


Private Sub StopButton_Click()
    Call DoStop
End Sub


Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    StopButton.Enabled = False
    StartButton.Enabled = True
    Slider.Enabled = False
    DevicesBox.Enabled = True
End Sub


Private Sub Visualize()
    Visualizing = True
    
    'These are all static just because they can.
    Static X As Long
    Static Wave As WaveHdr
    Static InData(0 To NumSamples - 1) As Integer
    Static OutData(0 To NumSamples - 1) As Single
    
    With ScopeBuff 'Save some time referencing it...
    
        Do
            Wave.lpData = VarPtr(InData(0))
            Wave.dwBufferLength = NumSamples
            Wave.dwFlags = 0
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
            
            Do
                'Just wait for the blocks to be done or the device to close
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            If DevHandle = 0 Then Exit Do 'Cut out if the device is closed
            
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
              
            Call FFTAudio(InData, OutData)
            
            .Cls
            .CurrentX = 0
            .CurrentY = ScopeHeight
        
            For X = 0 To 255
                .CurrentY = ScopeHeight
                .CurrentX = X
                
                'I average two elements here because it gives a smoother appearance.
                ScopeBuff.Line Step(0, 0)-(X, ScopeHeight - (Sqr(Abs(OutData(X * 2) \ Divisor)) + Sqr(Abs(OutData(X * 2 + 1) \ Divisor))))
            Next
            
            Scope.Picture = .Image 'Display the double-buffer
            DoEvents
        
        Loop While DevHandle <> 0
    
    End With
    
    Visualizing = False
End Sub
