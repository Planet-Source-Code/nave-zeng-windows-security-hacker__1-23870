VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Security Hacker - coded by Navarchy"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton focuser 
      Height          =   195
      Left            =   2.40000e5
      TabIndex        =   45
      Top             =   120
      Width           =   135
   End
   Begin VB.CommandButton about 
      BackColor       =   &H00FF0000&
      Caption         =   "About"
      Height          =   255
      Left            =   3360
      MaskColor       =   &H8000000A&
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "\/\/Information About This Program\/\/"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton openregedit 
      Caption         =   "RegEdit"
      Height          =   255
      Left            =   4080
      TabIndex        =   43
      ToolTipText     =   "Open the Windows Registry Editing Tool"
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton opendos 
      Caption         =   "DOS"
      Height          =   255
      Left            =   2520
      TabIndex        =   42
      ToolTipText     =   "Open a Command Prompt"
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   480
      Top             =   4440
   End
   Begin VB.Frame Frame5 
      Caption         =   "Logon Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   40
      Top             =   3360
      Width           =   2415
      Begin VB.TextBox logonmessage 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "This is the Message That Will Appear When a User Logs Onto the System"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "DOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   37
      Top             =   2280
      Width           =   2415
      Begin VB.CheckBox winoldapp 
         Caption         =   "NoRealMode"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Disable Single-Mode MS-DOS"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox winoldapp 
         Caption         =   "Disabled"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Disable MS-DOS Prompt"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Network"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5040
      TabIndex        =   30
      Top             =   240
      Width           =   2415
      Begin VB.CheckBox network 
         Caption         =   "NoPrintSharing"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Disables Print Sharing Controls"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox network 
         Caption         =   "NoFileSharingControl"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "Disables File Sharing Controls"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox network 
         Caption         =   "NoNetSetupSecurityPage"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Hides the Access Control Page"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox network 
         Caption         =   "NoNetSetupIDPage"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Hides the Identification Page"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CheckBox network 
         Caption         =   "NoNelSetup"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Hides the Network Option in the Control Panel"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox network 
         Caption         =   "NoNetSetupSecurityPage"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Hides Network Security Page"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2520
      TabIndex        =   16
      Top             =   600
      Width           =   2415
      Begin VB.CheckBox system 
         Caption         =   "NoVirtMemPage"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Hides Virtual Memory Button"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoFileSysPage"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Hides File System Button"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoConfigPage"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Hides Hardware Profiles Page"
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoDevMgrPage"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Hides Device Manager Page"
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoProfilePage"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Hides User Profiles Page"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoAdminPage"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Hides Remote Administration Page"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoPwdPage"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Hides Password Change Page"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoSecCPL"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Disables Password Control Panel"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoDispSettingsPage"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Hides Settings Page"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoDispAppearancePage"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Hides Appearance Page"
         Top             =   1080
         Width           =   2145
      End
      Begin VB.CheckBox system 
         Caption         =   "NoDispScrsavPage"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Hides Screen Saver Page"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox system 
         Caption         =   "NoDispBackgroundPage"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Hides Background Page"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox system 
         Caption         =   "NODispCPL"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Hides Control Panel"
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Explorer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox explorer 
         Caption         =   "Nolnternetlcon"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Removes the Internet (system folder) Icon From the Desktop"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "ClearRecentDocsOnExit"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Clears the Recent Documents System Folder on Shutdown"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoRecentDocsHistory"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Removes Recent Document System Folder From the Start Menu (IE 4 and above)"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "DisableRegistryTools"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Disable Registry Editing Tools"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoSaveSettings"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Don't Save Settings on Shutdown"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoClose"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Prevents the User From Normally Shutting Down Windows"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoDesktop"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Hides All Items From the Desktop"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoNetHood"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Hides theNetwork Neighborhood Icon From the Desktop"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoDrives"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Hides All of the Drives in My Computer"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoFind"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Removes the Find Tool (Start >Find)"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoSetTaskbar"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Removes Taskbar System Folder From the Settings Option in the Start Menu"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoSetFolders"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Removes Folders From the Settings Option in the Start Menu (Control Panel, Printers, Taskbar)"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoRun "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Disables or Hides the Run Command"
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoAddPrinter"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Disables Addition of New Printers"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox explorer 
         Caption         =   "NoDeletePrinter"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Disables Deletion of Already Installed Printers"
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''
'' Project Name: Windows Security Hacker  ''
''''''''''''''''''''''''''''''''''''''''''''
'' Description: Edits and monitors a      ''
''    number of registry keys that are    ''
''    used by Windows to store security   ''
''    information.                        ''
''''''''''''''''''''''''''''''''''''''''''''
'' Coder: Evan Sangaline                  ''
''    AKA Navarchy or Nave Zeng           ''
''''''''''''''''''''''''''''''''''''''''''''
'' Date: 6-7-01                           ''
''''''''''''''''''''''''''''''''''''''''''''
'' Code Status: This code is copyrighted  ''
''    and can only be distributed if      ''
''    absolutely no code is edited in any ''
''    way.                                ''
''''''''''''''''''''''''''''''''''''''''''''




' I always put this in my projects because
' it helps prevent errors.
Option Explicit

Private Sub about_Click()
    ' Set Timer2 to 0 so that the button
    ' doesn't blick any more
    Timer2.Interval = 0
    
    ' Set the button color to gray so that
    ' it looks normal
    about.BackColor = &H8000000A
    
    ' Get rid of the arrows in the tooltip text
    about.ToolTipText = "Information About This Program"
    
    ' I thought it looked bad when about
    ' had the focus so I created a button
    ' that you can't see to take the focus
    ' after opendos is clicked.
    focuser.SetFocus
    
    ' This is a message box that explains a little
    ' abou the program. I know its hard to read on
    ' one line, so sorry. Please don't change this
    ' because I wrote the program and want my name
    ' on it.
    MsgBox "                                                          Windows Security Hacker" & vbCrLf & "                                                             coded by Navarchy" & vbCrLf & vbCrLf & "This program can be used by either administrators who need to limmit access or users who desire more access. There are many functions inside of this program so a few of them may not work on some Windows operating systems, but the majority of them work on all Windows operating systems. " & vbCrLf & "                                                                        Peace " & vbCrLf & "                                                                      Navarchy", vbSystemModal, ""
End Sub

Private Sub explorer_Click(Index As Integer)
    ' Create the registry path if it doesn't already exist
    CreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    
    ' The caption of each of the checkboxes
    ' is the same as the corresponding registry
    ' key. The key value set here is used by
    ' Windows to store the corresponding security
    ' information. The last part just converts
    ' the value of the checkbox to a long for
    ' the function.
    SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", explorer(Index).Caption, CLng(explorer(Index).Value)
End Sub

Sub Updater()
  Dim a As Integer
    ' This goes through each of the checkboxes
    ' in the explorer array and sets each checkboxes
    ' to the value recieved from the registry.
    For a = 0 To explorer.Count - 1
        explorer(a).Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", explorer(a).Caption, 0)
    Next a
    ' Set a to 0 so that I can use the same
    ' variable again
    a = 0
    
    ' This goes through each of the checkboxes
    ' in the system array and sets each checkboxes
    ' to the value recieved from the registry.
    For a = 0 To system.Count - 1
        system(a).Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", system(a).Caption, 0)
    Next a
    ' Set a to 0 so that I can use the same
    ' variable again
    a = 0
    
    ' This goes through each of the checkboxes
    ' in the network array and sets each checkboxes
    ' to the value recieved from the registry.
    For a = 0 To network.Count - 1
        network(a).Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", network(a).Caption, 0)
    Next a
    
    ' There are only two things in the winoldapp
    ' array, so it takes less code to just write
    ' them individually.
    winoldapp(0).Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", winoldapp(0).Caption, 0)
    winoldapp(1).Value = GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", winoldapp(1).Caption, 0)
    
    ' This retrieves the logon message string
    ' from the registry and puts the string
    ' into logonmessage.Text
    logonmessage.Text = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\WinLogon", "LegalNoticeCaption", "")
End Sub

Private Sub Form_Load()
    
  Dim sBuffer As String
  Dim lSize As Long
  Dim getusername As String
    ' Space for dll parameters
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    ' Get the username
    Call GetUserNameAPI(sBuffer, lSize)

    If lSize > 0 Then
        ' Remove empty spaces
        getusername = Left$(sBuffer, lSize)
      Else
        ' Return empty if no user is found
        getusername = vbNullString
    End If
    ' Add the username to the caption of the form
    Me.Caption = Me.Caption & " - current user: " & getusername
    
    ' Center the form
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    ' Call the sub that sets the checkboxes
    ' and textbox to the values in the registry
    Updater
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Set the timers to 0 to avoid delay
    Timer1.Interval = 0
    Timer2.Interval = 0
    
    ' Unload the form
    Unload Me
    
    ' End the project
    End
End Sub

Private Sub network_Click(Index As Integer)
    ' Create the registry path if it doesn't already exist
    CreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network"
    
    ' The caption of each of the checkboxes
    ' is the same as the corresponding registry
    ' key. The key value set here is used by
    ' Windows to store the corresponding security
    ' information. The last part just converts
    ' the value of the checkbox to a long for
    ' the function.
    SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", network(Index).Caption, CLng(network(Index).Value)
End Sub

Private Sub opendos_Click()
    ' This ends up calling winoldapp_Click
    ' and enabling access to the DOS command prompt
    winoldapp(0).Value = 0
    
    ' I thought it looked bad when opendos
    ' had the focus so I created a button
    ' that you can't see to take the focus
    ' after opendos is clicked.
    focuser.SetFocus
    
    ' This executes command.com (DOS) in
    ' its normal state with the focus.
    Shell "command.com", vbNormalFocus
End Sub

Private Sub openregedit_Click()
    ' This ends up calling explorer_Click
    ' and enabling access to the Windows
    ' registry editing tools.
    explorer(11).Value = False
    
    ' I thought it looked bad when openregedit
    ' had the focus so I created a button
    ' that you can't see to take the focus
    ' after openregedit is clicked.
    focuser.SetFocus
    
    ' This executes regedit.exe (the Windows registry editor) in
    ' its normal state with the focus.
    Shell "regedit.exe", vbNormalFocus
End Sub

Private Sub system_Click(Index As Integer)
    ' Create the registry path if it doesn't already exist
    CreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    
    ' The caption of each of the checkboxes
    ' is the same as the corresponding registry
    ' key. The key value set here is used by
    ' Windows to store the corresponding security
    ' information. The last part just converts
    ' the value of the checkbox to a long for
    ' the function.
    SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", system(Index).Caption, CLng(system(Index).Value)
End Sub

Private Sub logonmessage_Change()
    ' Create the registry path if it doesn't already exist
    CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\WinLogon"
    
    ' The caption of each of the checkboxes
    ' is the same as the corresponding registry
    ' key. The key value set here is used by
    ' Windows to store the corresponding security
    ' information. The last part just converts
    ' the value of the checkbox to a long for
    ' the function.
    SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\WinLogon", "LegalNoticeCaption", logonmessage.Text
End Sub

Private Sub Timer1_Timer()
    ' Call the sub that sets the checkboxes
    ' and textbox to the values in the registry
    Updater
End Sub

Private Sub Timer2_Timer()
    ' This sub just draws attention to the
    ' about button so that people click it
    ' more often.
    
    ' If the about button is red then
    If about.BackColor = &HFF& Then
        ' make the tool tip text arrows point down
        about.ToolTipText = "\/\/Information About This Program\/\/"
        ' and make it blue
        about.BackColor = &HFF0000
      ' If it is blue then
      Else
        ' make it red
        about.BackColor = &HFF&
        ' and make the tool tip text arrows point up
        about.ToolTipText = "/\/\Information About This Program/\/\"
    End If
End Sub

Private Sub WinOldApp_Click(Index As Integer)
    ' Create the registry path if it doesn't already exist
    CreateKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
    
    ' The caption of each of the checkboxes
    ' is the same as the corresponding registry
    ' key. The key value set here is used by
    ' Windows to store the corresponding security
    ' information. The last part just converts
    ' the value of the checkbox to a long for
    ' the function.
    SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", winoldapp(Index).Caption, CLng(winoldapp(Index).Value)
End Sub
