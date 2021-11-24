VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{9EB67EA9-DBFB-11D2-92B7-00C0DFE9A30F}#3.0#0"; "BWMIXER1BVB6.OCX"
Begin VB.Form Form2 
   Caption         =   "BinaryWork Mixer OCX 1.0 - Sample project"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form2"
   ScaleHeight     =   3045
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Mute"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Mute"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mute"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin ComctlLib.Slider Slider4 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   327682
      Max             =   100
   End
   Begin ComctlLib.Slider Slider3 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   327682
      Max             =   100
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   327682
      Max             =   100
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mute"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   327682
      Max             =   100
   End
   Begin BwMixerOCX.BWMixer BWMixer1 
      Left            =   4200
      Top             =   2160
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin VB.Label Label4 
      Caption         =   "Wave out"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Midi out"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Line In"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Master"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MasterId(0 To 1) As Long
Dim WaveId(0 To 1) As Long
Dim MidiId(0 To 1) As Long
Dim LineInId(0 To 1) As Long
Dim MasterMuteId(0 To 1) As Long
Dim WaveMuteId(0 To 1) As Long
Dim MidiMuteId(0 To 1) As Long
Dim LineInMuteId(0 To 1) As Long


Private Sub BWMixer1_ReturnDevicesAvailable(DeviceName As String, ShortDeviceName As String, KnownDevice As String, DeviceId As Long, DeviceType As Long, MaximumValue As Long)

Select Case KnownDevice

Case "WaveOutput"
WaveId(0) = DeviceId
WaveId(1) = DeviceType

Case "MidiOutput"
MidiId(0) = DeviceId
MidiId(1) = DeviceType

Case "LineInOutput"
LineInId(0) = DeviceId
LineInId(1) = DeviceType

Case "MasterVolume"
MasterId(0) = DeviceId
MasterId(1) = DeviceType

Case "WaveOutMute"
WaveMuteId(0) = DeviceId
WaveMuteId(1) = DeviceType

Case "MidiOutMute"
MidiMuteId(0) = DeviceId
MidiMuteId(1) = DeviceType

Case "LineInMute"
LineInMuteId(0) = DeviceId
LineInMuteId(1) = DeviceType

Case "MasterVolumeMute"
MasterMuteId(0) = DeviceId
MasterMuteId(1) = DeviceType

End Select

End Sub



Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ret As Long

ret = BWMixer1.SetValueByType(MasterMuteId(0), MasterMuteId(1), Check1.Value)

End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ret As Long
ret = BWMixer1.SetValueByType(MidiMuteId(0), MidiMuteId(1), Check2.Value)
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ret As Long

ret = BWMixer1.SetValueByType(LineInMuteId(0), LineInMuteId(1), Check3.Value)
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ret As Long

ret = BWMixer1.SetValueByType(WaveMuteId(0), WaveMuteId(1), Check4.Value)
End Sub

Private Sub Command1_Click()
Form2.Caption = " Ola"
End Sub

Private Sub Form_Activate()

BWMixer1.InitializeMixer
Timer1.Enabled = True

End Sub




Private Sub Slider1_Scroll()

Dim ret As Long
ret = BWMixer1.SetValueByType(MasterId(0), MasterId(1), Slider1.Value)

End Sub

Private Sub Slider2_Scroll()

Dim ret As Long
ret = BWMixer1.SetValueByType(LineInId(0), LineInId(1), Slider2.Value)

End Sub

Private Sub Slider3_Scroll()
Dim ret As Long

ret = BWMixer1.SetValueByType(MidiId(0), MidiId(1), Slider3.Value)
End Sub

Private Sub Slider4_Scroll()
Dim ret As Long

ret = BWMixer1.SetValueByType(WaveId(0), WaveId(1), Slider4.Value)
End Sub

Private Sub Timer1_Timer()

If BWMixer1.IsInitialized = False Then Timer1.Enabled = False: Exit Sub


If BWMixer1.IsUpdatingVolumes = True Then Exit Sub

If MasterId(0) <> 0 Then Slider1.Value = BWMixer1.GetValueByType(MasterId(0), MasterId(1))

If MidiId(0) <> 0 Then Slider3.Value = BWMixer1.GetValueByType(MidiId(0), MidiId(1))

If LineInId(0) <> 0 Then Slider2.Value = BWMixer1.GetValueByType(LineInId(0), LineInId(1))

If WaveId(0) <> 0 Then Slider4.Value = BWMixer1.GetValueByType(WaveId(0), WaveId(1))

If MasterMuteId(0) <> 0 Then Check1.Value = BWMixer1.GetValueByType(MasterMuteId(0), MasterMuteId(1))

If MidiMuteId(0) <> 0 Then Check2.Value = BWMixer1.GetValueByType(MidiMuteId(0), MidiMuteId(1))

If LineInMuteId(0) <> 0 Then Check3.Value = BWMixer1.GetValueByType(LineInMuteId(0), LineInMuteId(1))

If WaveMuteId(0) <> 0 Then Check4.Value = BWMixer1.GetValueByType(WaveMuteId(0), WaveMuteId(1))

End Sub
