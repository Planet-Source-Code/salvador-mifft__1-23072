VERSION 5.00
Begin VB.Form HzView 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spectrum Viewer"
   ClientHeight    =   3912
   ClientLeft      =   1440
   ClientTop       =   2748
   ClientWidth     =   6576
   DrawMode        =   8  'Xor Pen
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   548
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox HyPiano 
      Height          =   372
      Left            =   2760
      ScaleHeight     =   324
      ScaleWidth      =   924
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Timer QuitTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   3720
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2856
      Left            =   120
      ScaleHeight     =   238
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   521.377
      TabIndex        =   2
      Top             =   360
      Width           =   6672
   End
   Begin VB.CommandButton StopButton 
      BackColor       =   &H0000C000&
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   336
      Left            =   1800
      TabIndex        =   3
      Top             =   3480
      Width           =   984
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   288
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   3108
   End
   Begin VB.CommandButton StartButton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Start"
      Height          =   336
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   984
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   336
      Left            =   5520
      ScaleHeight     =   28
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   98
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1176
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "22,050 Hz"
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   5640
      TabIndex        =   5
      Top             =   0
      Width           =   972
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu mpiano 
         Caption         =   "Piano"
      End
   End
End
Attribute VB_Name = "HzView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
'   An oscilloscope and Audio Spectrum "Viewer" with an hyper piano to help interpret the output.
'
'   I 've been trying to understand the output of the Fast Fourier Transform for a proyect I'm working on
'in which I plan to use the dominant frequency to associate sound with color in an analogic way (low frequency
'colors with low frequency sounds, etc.)
'    I made this example, based on Murphy McCauley's Deeth Spectrum Analyzer v1.0 changing only the
'portions of code that I needed for my purpouses. (By the way, someone uploaded the original program
'some months ago, changing only the name of the author).
'    The program is oriented to help understand how digital audio is recorded and how you can use it for
'an aplication (your own CD or MP3 player) using the FFT.
'   The program will graph only de peak frequency (the loudest) in each sample of 1024 (every  0.0232 sec.)
'drawing lines in the position given by that frequency, using some tricks to be able to represent 22,050 positions
'in a 512 width picturebox.
' Remember to control the volume of Recording if you don't see any effects. The program output
' may change depending on the speed of your machine
'  Hope you like the result and find the code useful.
'  If you want to use the original code visit http://www.fullspectrum.com/deeth/
'----------------------------------------------------------------------

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
Dim maxvol As Long, Hz As Long, oscila As Long
Dim HzColor As Long, xMax As Integer, HzTip As Long


Sub InitDevices()
    'Fill the DevicesBox box with all the compatible audio input devices
    'Bail if there are none.
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        If Caps.Formats And WAVE_FORMAT_4M16 Then '16-bit mono devices
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
    HyPiano.Picture = LoadPicture(App.Path & "\hypiano.bmp")
    Call InitDevices 'Fill the DevicesBox
    Call DoReverse   'Pre-calculate these
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
Private Sub mpiano_Click()
If mpiano.Checked = False Then mpiano.Checked = True Else mpiano.Checked = False
End Sub
Private Sub QuitTimer_Timer()
    Unload Me
End Sub
Private Sub Scope_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Y >= ScaleHeight / 3 And Y < ScaleHeight / 3 * 2 Then Scope.ToolTipText = "516.84 to 5,125.33Hz": Exit Sub
  If Y >= ScaleHeight / 3 * 2 Then Scope.ToolTipText = "43.07 to 473.77 Hz": Exit Sub
  Scope.ToolTipText = "5,168 to 22K Hz"
End Sub
Private Sub StartButton_Click()
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = 44100
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    maxvol = waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Call waveInStart(DevHandle)
    StopButton.Caption = "&Stop"
    StopButton.Enabled = True
    StartButton.Enabled = False
    DevicesBox.Enabled = False
    Call Visualize
End Sub
Private Sub StopButton_Click()
    Call DoStop
    If StopButton.Caption = "&Stop" Then
    StopButton.Caption = "Exit"
    Exit Sub
    End If
    QuitTimer = True
End Sub
Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    'StopButton.Enabled = False
    StartButton.Enabled = True
    DevicesBox.Enabled = True
End Sub
Private Sub Visualize()
'                    Original code to get the data and process the FFT
    Static X As Long
    Static Wave As WaveHdr
    Static InData(0 To NumSamples - 1) As Integer
    Static OutData(0 To NumSamples - 1) As Single
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
            
            
'                Code to represent 3 levels of graphic display to "compress" the range in 512 pixels
'                (decided to draw an hyperpiano... hyper because the frequency range of the piano
'                is limited for the 22KHz that we may visualize after the FFT. (The lower portion
'                of the "viewer" corresponds partially to the corresponding notes in a real piano
              ScopeBuff.BackColor = vbBlack
              oscila = vbGreen
              ' decided to capture the screen and use a bmp instead (for faster performance)
              If mpiano.Checked = True Then ScopeBuff.Picture = HyPiano.Picture: oscila = vbBlack Else ScopeBuff.Picture = LoadPicture()
   'If mpiano.Checked = True Then '       draw the hyperpiano
   '           oscila = vbBlack
   '           ScopeBuff.BackColor = StartButton.BackColor
   '           ' draw the white keys (range 43.07 to 516.84Hz
   '         For X = 0 To Scope.ScaleWidth + 50 Step 48
   '           ScopeBuff.Line (X, (ScopeBuff.ScaleHeight / 3) * 2)-(X, ScopeBuff.ScaleHeight), vbBlack
   '         Next X
   '           ' draw the black keys (range 516hz to 5,125KHz
   '         For X = 0 To Scope.ScaleWidth + 50 Step 24
   '           ScopeBuff.Line (X, ScopeBuff.ScaleHeight / 3)-(X, (ScopeBuff.ScaleHeight)), vbBlack  'RGB(100, 110, 120)
   '         Next X
   '         ScopeBuff.DrawWidth = 5
   '           ' but don't draw the semitones
   '         For X = 0 To Scope.ScaleWidth + 50 Step 24
   '           If X = 24 * 3 Then GoTo semi
   '           If X = 24 * 7 Then GoTo semi
   '           If X = 24 * 10 Then GoTo semi
   '           If X = 24 * 14 Then GoTo semi
   '           If X = 24 * 17 Then GoTo semi
   '           If X = 24 * 21 Then GoTo semi
   '           ScopeBuff.Line (X - 1, ScopeBuff.ScaleHeight / 3 + 1)-(X - 1, (ScopeBuff.ScaleHeight / 3 * 2)), vbBlack 'RGB(100, 110, 120)
'semi:
   '         Next X
            ' draw the third range (5168 to 22050Hz)
   '         ScopeBuff.DrawWidth = 1
    '        For X = 1 To Scope.ScaleWidth + 50 Step 12
    '          ScopeBuff.Line (X, 0)-(X, (ScopeBuff.ScaleHeight / 2)), vbBlack 'RGB(100, 110, 120)
    '        Next X
    '        ScopeBuff.Line (0, ScopeBuff.ScaleHeight / 3)-(ScopeBuff.ScaleWidth, ScopeBuff.ScaleHeight / 3), vbBlack
            Dim c As Double, LowMidHig
  'End If
            
'                    My Code to determine the single dominant frequency per sample (1024)
            For X = 1 To 511
               ScopeBuff.DrawWidth = 1
               ScopeBuff.PSet (X, ScopeHeight / 3 - (InData(X) / 500)), oscila ' oscilloscope
               'ScopeBuff.PSet (X, ScopeHeight - OutData(X) / 500), oscila ' FFT out (just for reference)
            If Abs(OutData(X)) > maxvol Then
               maxvol = Abs(OutData(X))
               Hz = Int(44100 * X) / 1024
               Label1.Caption = Hz & " Hz"
               HzColor = vbRed '+ Hz
               LowMidHig = ScopeHeight
               '  to draw 1st range (red lines almost correspond to the actual freq in a piano)
               '  (the first white key corresponds to A 55Hz ... the line shows 43.07Hz)
               If X < 11 Then Hz = Hz - 22 'try to center line on piano key
               '  second range the lines will be drawn in their relative position as freq
               If X > 11 Then Hz = Hz / 10: HzColor = StopButton.BackColor + X: LowMidHig = (ScopeHeight / 3) * 2
               '   third range ... quite high ... the lines are shown sequentially as they appear
               '   in the buffer and the volume is doubled (normally those high sounds are harmonics) and
               '   they rarely will be the dominant frquency.
               If X > 119 Then Hz = X - 119: HzColor = vbYellow: LowMidHig = (ScopeHeight / 3): OutData(X) = OutData(X) * 2
               xMax = X
            End If
            Next
            X = xMax
                   c = 0.5 * (1 - Cos(X * 2 * 3.1416 / 512)) 'Hanning Window
                   OutData(X) = c * OutData(X)
                '  use color asociated with the value of frequency
                If mpiano.Checked = False Then HzColor = (Int(44100 * X) / 1024) * 10000
                ScopeBuff.DrawWidth = 5
                If X > 11 Then ScopeBuff.DrawWidth = 3
                If X > 119 Then ScopeBuff.DrawWidth = 4
                ScopeBuff.Line (Hz, LowMidHig)-(Hz, LowMidHig - ((Abs(OutData(X)) / 10))), HzColor
                ScopeBuff.DrawWidth = 1
                maxvol = 0
                Scope.Picture = ScopeBuff.Image 'Display the double-buffer
                'Sleep 60
                ScopeBuff.Cls
            DoEvents
        Loop While DevHandle <> 0
End Sub
Private Sub prueba(mioutdata)

End Sub

