VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{9EB67EA9-DBFB-11D2-92B7-00C0DFE9A30F}#3.0#0"; "BWMIXER1BVB6.OCX"
Begin VB.Form Form1 
   Caption         =   "BinaryWork Mixer OCX - Sample Project"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3600
      TabIndex        =   9
      Top             =   4680
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   1560
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   327682
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   50
      TabIndex        =   1
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5741
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin BwMixerOCX.BWMixer BWMixer1 
      Left            =   6120
      Top             =   3600
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin VB.Label Label3 
      Caption         =   "Mixers available in the system"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   5055
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   5055
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4815
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Click in the list to select the device to be controled"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4815
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Definition of the variables that are required to the mixer

Dim DeviceSelectionName As String  'the name of the device returned by the mixer
Dim DeviceSelection As Long  'The DeviceID that is required to get and set the value
Dim DeviceTypeSelected As Long 'The DeviceType that is required also in order to get and set the
Dim DeviceMaximumValue As Long 'The maximum value of the actual device selected , this is required also

Dim NewValue As Long

Private Sub About_Click()
BWMixer1.About
End Sub

'Flags to avoid the refresh of the form if the value have not changed since the last update

'This event return any error generated in the mixer , with description also
Private Sub BWMixer1_MixerError(ErrorCode As Long, errorDescription As String)
Form1.Caption = errorDescription

End Sub

'This event will return the mixer properites , this was called by InitialiZe mixer or GetMixerProperties
Private Sub BWMixer1_ReturnMixerProperties(MixerName As String, ManufacturerId As Integer, ProductId As Integer, DriverVersion As Long)

Label11.Caption = "Driver version : " & DriverVersion
Label9.Caption = "Product ID : " & ProductId
Label7.Caption = "Manufacturer ID : " & ManufacturerId
Label5.Caption = "Mixer description : " & MixerName

End Sub

'"This event will list any device detected by the Mixer , and here you can see
'the extended capabilities of our mixer , this will list any information that is required in order to access
'and change any device available
'This was called by InitializeMixer

Private Sub BWMixer1_ReturnDevicesAvailable(DeviceName As String, ShortDeviceName As String, KnownDevice As String, DeviceId As Long, DeviceType As Long, MaximumValue As Long)

Set minhalista = ListView1.ListItems.Add(1, , DeviceName)
ListView1.ListItems.Item(1).SubItems(1) = DeviceId
ListView1.ListItems.Item(1).SubItems(2) = DeviceType
ListView1.ListItems.Item(1).SubItems(3) = KnownDevice
ListView1.ListItems.Item(1).SubItems(4) = MaximumValue
End Sub

Private Sub BWMixer1_ReturnMixersInTheSystem(MixerName As String, ManufacturerId As Integer, ProductId As Integer, DriverVersion As Long)
'it will fill the listbox with the name of the mixers in the system
List1.AddItem MixerName

End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ret As Long
ret = BWMixer1.SetValueByType(DeviceSelection, DeviceTypeSelected, Check1.Value)

End Sub



Private Sub Command1_Click()


End Sub

Private Sub Form_Activate()

'This will create the definition of the listview1

Dim clmX As ColumnHeader
    Set clmX = ListView1.ColumnHeaders. _
    Add(, , "Device Name", ListView1.Width / 4)
Set clmX = ListView1.ColumnHeaders. _
    Add(, , "Device ID", ListView1.Width / 12)
    Set clmX = ListView1.ColumnHeaders. _
    Add(, , "Device Type", ListView1.Width / 8)
    Set clmX = ListView1.ColumnHeaders. _
    Add(, , "Known Device", ListView1.Width / 6)
    Set clmX = ListView1.ColumnHeaders. _
    Add(, , "Max Value", ListView1.Width / 8)

ListView1.View = lvwReport

'This will Intialize the mixer in the first activation of the form . without it the mixer will not work or update

If BWMixer1.IsInitialized = False Then
BWMixer1.InitializeMixer
End If

'this will raise the events with the information about the mixers available and their descriptions

BWMixer1.GetMixersAvailable
BWMixer1.GetMixerProperties

End Sub


Private Sub ListView1_Click()

If BWMixer1.IsInitialized = False Then Exit Sub 'This will verify if the mixer is initialized

DeviceSelectionName = ListView1.SelectedItem
DeviceSelection = ListView1.SelectedItem.SubItems(1)
DeviceTypeSelected = ListView1.SelectedItem.SubItems(2)
DeviceMaximumValue = ListView1.SelectedItem.SubItems(4)

If DeviceMaximumValue > 1 Then  'it will verify if the value need to be expressed in a checkbox or a slider
        
        'the folowing flags will avoid the update of the objects if the value is the same of the last update of the mixer

        'without it the slider or the checkbox will appear strange , because the update affect the entire form

        'then it will be updated only if it is needed

            If Slider1.Enabled = False Then Slider1.Enabled = True  'If the slider isnot enabled then enable it
            
            If Label2.Caption <> DeviceSelectionName Then Label2.Caption = DeviceSelectionName  'update the caption if it is needed
            
            If Check1.Enabled = True Then Check1.Enabled = False   'disable the checkbox if is is enable
            
            If Check1.Caption <> "" Then Check1.Caption = ""
            
            If Slider1.Min <> 0 Then Slider1.Min = 0
            
            If Slider1.Max <> DeviceMaximumValue Then Slider1.Max = DeviceMaximumValue
            
            NewValue = BWMixer1.GetValueByType(DeviceSelection, DeviceTypeSelected) 'get the new value of the referred mixer to be compared
            
            If Slider1.Value <> NewValue Then Slider1.Value = NewValue 'update the slider if it is required

Else

            If Check1.Enabled = False Then Check1.Enabled = True
               
            If Label2.Caption <> "" Then Label2.Caption = ""
            
            If Slider1.Enabled = True Then Slider1.Enabled = False
            
            If Slider1.Value <> 0 Then Slider1.Value = 0
            
            If Check1.Caption <> DeviceSelectionName Then Check1.Caption = DeviceSelectionName
            
            NewValue = BWMixer1.GetValueByType(DeviceSelection, DeviceTypeSelected)
            
            If Check1.Value <> NewValue Then Check1.Value = NewValue
            
End If

If Timer1.Enabled = False Then Timer1.Enabled = True  'it will start the timer if it isnot already running

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

If BWMixer1.IsInitialized = False Then Exit Sub 'This will verify if the mixer is initialized

DeviceSelectionName = ListView1.SelectedItem
DeviceSelection = ListView1.SelectedItem.SubItems(1)
DeviceTypeSelected = ListView1.SelectedItem.SubItems(2)
DeviceMaximumValue = ListView1.SelectedItem.SubItems(4)

If DeviceMaximumValue > 1 Then  'it will verify if the value need to be expressed in a checkbox or a slider
        
        'the folowing flags will avoid the update of the objects if the value is the same of the last update of the mixer

        'without it the slider or the checkbox will appear strange , because the update affect the entire form

        'then it will be updated only if it is needed

            If Slider1.Enabled = False Then Slider1.Enabled = True  'If the slider isnot enabled then enable it
            
            If Label2.Caption <> DeviceSelectionName Then Label2.Caption = DeviceSelectionName  'update the caption if it is needed
            
            If Check1.Enabled = True Then Check1.Enabled = False   'disable the checkbox if is is enable
            
            If Check1.Caption <> "" Then Check1.Caption = ""
            
            If Slider1.Min <> 0 Then Slider1.Min = 0
            
            If Slider1.Max <> DeviceMaximumValue Then Slider1.Max = DeviceMaximumValue
            
            NewValue = BWMixer1.GetValueByType(DeviceSelection, DeviceTypeSelected) 'get the new value of the referred mixer to be compared
            
            If Slider1.Value <> NewValue Then Slider1.Value = NewValue 'update the slider if it is required

Else

            If Check1.Enabled = False Then Check1.Enabled = True
               
            If Label2.Caption <> "" Then Label2.Caption = ""
            
            If Slider1.Enabled = True Then Slider1.Enabled = False
            
            If Slider1.Value <> 0 Then Slider1.Value = 0
            
            If Check1.Caption <> DeviceSelectionName Then Check1.Caption = DeviceSelectionName
            
            NewValue = BWMixer1.GetValueByType(DeviceSelection, DeviceTypeSelected)
            
            If Check1.Value <> NewValue Then Check1.Value = NewValue
            
End If

If Timer1.Enabled = False Then Timer1.Enabled = True  'it will start the timer if it isnot already running

End Sub

Private Sub Slider1_Scroll()

Dim ret As Long
ret = BWMixer1.SetValueByType(DeviceSelection, DeviceTypeSelected, Slider1.Value)

End Sub

Private Sub Timer1_Timer()

If BWMixer1.IsInitialized = False Then Exit Sub  'it will verify if the Mixer was initialized
If BWMixer1.IsUpdatingVolumes = True Then Exit Sub      'it will verify if the mixer is already making some changes in the system mixer , if so , then wait until the changes finished

If DeviceMaximumValue > 1 Then  'it will verify if the value need to be expressed in a checkbox or a slider
        
        'the folowing flags will avoid the update of the objects if the value is the same of the last update of the mixer

        'without it the slider or the checkbox will appear strange , because the update affect the entire form

        'then it will be updated only if it is needed

            If Slider1.Enabled = False Then Slider1.Enabled = True  'If the slider isnot enabled then enable it
            
            If Label2.Caption <> DeviceSelectionName Then Label2.Caption = DeviceSelectionName  'update the caption if it is needed
            
            If Check1.Enabled = True Then Check1.Enabled = False   'disable the checkbox if is is enable
            
            If Check1.Caption <> "" Then Check1.Caption = ""
            
            If Slider1.Min <> 0 Then Slider1.Min = 0
            
            If Slider1.Max <> DeviceMaximumValue Then Slider1.Max = DeviceMaximumValue
            
            NewValue = BWMixer1.GetValueByType(DeviceSelection, DeviceTypeSelected) 'get the new value of the referred mixer to be compared
            
            If Slider1.Value <> NewValue Then Slider1.Value = NewValue 'update the slider if it is required

Else

            If Check1.Enabled = False Then Check1.Enabled = True
               
            If Label2.Caption <> "" Then Label2.Caption = ""
            
            If Slider1.Enabled = True Then Slider1.Enabled = False
            
            If Slider1.Value <> 0 Then Slider1.Value = 0
            
            If Check1.Caption <> DeviceSelectionName Then Check1.Caption = DeviceSelectionName
            
            NewValue = BWMixer1.GetValueByType(DeviceSelection, DeviceTypeSelected)
            
            If Check1.Value <> NewValue Then Check1.Value = NewValue
            
End If

End Sub


