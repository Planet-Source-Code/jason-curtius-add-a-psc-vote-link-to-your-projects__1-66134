VERSION 5.00
Begin VB.UserControl vLink 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   465
   ScaleWidth      =   2295
   Windowless      =   -1  'True
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Please click here to vote."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "vLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Add a vote link by Jason Curius
'
'The purpose of this user control is to simply allow you
'to place a vote/comment link on your projects at PSC.
'Simply add this to any of your projects and configure
'the propities to suit before uploading your project.
'
'This wont get you more votes, but it will make it
'far more convenient for those who do vote.
'
'The way this works is by locating the @PSC_ReadMe.txt
'file that is automaticly added to all projects at PSC.
'then finding the url within this file and linking to it.
'
'From this code you can see how to :
' - make costom links to web pages
' - launch/run most kinds of files
' - scan folders and subfolders
' - read a text file line by line
' - create propities for user controls
' - return True/False from an Integer
' - return an Integer from True/False
' - and some other basics
'

Option Explicit ' this can save you alot of pain. All it means is all Variables must be declared.
'Reference: OLE Atomation
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ShellExecute Lib Declaration allows us to launch
'almost any kind of file. In this case we will be
'using it to open a web page.(the vote/comment page for your project)
'
'set hand cursor Declarations
'from http://www.maxvb.com/post-5626.html
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'
'If you get a Compile Error see below:
Private hvrC As OLE_COLOR 'hover color
'OLE Atomation' IN Project\References MUST BE CHECKED.

Private oldC As OLE_COLOR 'forecolor after hover

Public Event Change() ' The Change event

Public ThankyouMessage As String 'text after link is clicked

Private newFont As Font 'newFont used in property bag
Private Sub ScanForPSC_ReadMe(ByVal ProjectPath As String)
'This is where we scan for files within our project folder
'using the FileSystemObject.
    
    Dim capLength As Variant 'length of Caption for our Link
    
    Dim LineCount As Variant ' count lines of text
    
    Dim txtFile As String 'string to hold text file name
    
    Dim PSCLineTxt As String 'string to hold a single line of text from textfile
    
    Dim fso As Object 'sets 'fso' as our reference to FileSystemObject(microsoft scripting runtime)
    
    Dim projectFolders As Object 'holds folder names
    
    Dim projectFiles As Object 'holds file names
    
    'Dim subDirectory As Folder 'To scan sub folders we don't need to in this case.
    On Error Resume Next ' ignores errors, you never know what we might find or more importantly not find.
    Set fso = CreateObject("scripting.filesystemobject")
    Set projectFolders = fso.GetFolder(ProjectPath) ' Our project folder

'look for @PSC_ReadMe.txt file within project folder
'by looking for a file that begins with this "@PSC_ReadMe_"
For Each projectFiles In projectFolders.Files
    If Len(projectFolders) > 3 Then 'just in case we're in a root directory we'll need to add a "\" into our path
        If Mid$(projectFiles.Name, 1, 12) = "@PSC_ReadMe_" Then txtFile = (projectFolders & "\" & projectFiles.Name)
        Else
        If Mid$(projectFiles.Name, 1, 12) = "@PSC_ReadMe_" Then txtFile = (projectFolders & projectFiles.Name)
    End If
Next
'when a file starting with "@PSC_ReadMe_" is found txtFile
'its full path is stored in this string "txtFile"
'
'******************************************************
'This part is not needed but I want to shows how to
'scan all sub folders in the current directory
'
'    For Each subDirectory In projectFolders.SubFolders
'        ShowAllFiles subDirectory.Path
'    Next
'******************************************************

'this next bit is ignored if @PSC_ReadMe.txt is found
If txtFile = "" Then 'not here, then look in parent folder
For Each projectFiles In projectFolders.ParentFolder.Files
    If Len(projectFolders) > 3 Then
        If Mid$(projectFiles.Name, 1, 12) = "@PSC_ReadMe_" Then txtFile = (projectFolders.ParentFolder & "\" & projectFiles.Name)
        Else
        If Mid$(projectFiles.Name, 1, 12) = "@PSC_ReadMe_" Then txtFile = (projectFolders.ParentFolder & projectFiles.Name)
    End If
Next

End If
'again this next bit is ignored if @PSC_ReadMe.txt is found
If txtFile = "" Then 'still not found, up 1 more folder for a last try
For Each projectFiles In projectFolders.ParentFolder.ParentFolder.Files
    If Len(projectFolders) > 3 Then
        If Mid$(projectFiles.Name, 1, 12) = "@PSC_ReadMe_" Then txtFile = (projectFolders.ParentFolder.ParentFolder & "\" & projectFiles.Name)
        Else
        If Mid$(projectFiles.Name, 1, 12) = "@PSC_ReadMe_" Then txtFile = (projectFolders.ParentFolder.ParentFolder & projectFiles.Name)
    End If
Next
End If
   'Set subDirectory = Nothing' only used to scan sub folders
    Set projectFiles = Nothing
    Set projectFolders = Nothing
    Set fso = Nothing
'now read the @PSC_ReadMe.txt, assuming it exists
    Open txtFile For Input As #1
'200 is how many lines to read before giving up
'most likely overkill but it's good to be safe
        Do While LineCount < 200 And Mid$(PSCLineTxt, 1, 57) <> "You can view comments on this code/and or vote on it at: "
        LineCount = LineCount + 1
        Line Input #1, PSCLineTxt 'This reads line by line
        Loop
    Close #1
'got the line of text containing the url
'now to get just the url from the PSCLineTxt
PSCLineTxt = Mid$(PSCLineTxt, 58, Len(PSCLineTxt))
capLength = Len(Label1.Caption) 'get length of Link text
If PSCLineTxt <> "" Then

If Label1.Caption = ThankyouMessage Then Exit Sub 'has already voted then exit
'open the web page
If ShellExecute(hwnd, "open", PSCLineTxt, "", "", 1) <= 32 Then Call ShellExecute(hwnd, "open", "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " + PSCLineTxt, "%windir%\system32", 1)
Label1.Caption = ThankyouMessage 'show thankyou message
Else
Label1.Caption = "URL Not Found." 'Error can't find file or url
End If
'resize user control only if more space is needed for text
If capLength < Len(Label1.Caption) Then UserControl_Resize
End Sub

Private Sub Label1_Click()
ScanForPSC_ReadMe (App.Path) 'look for @PSC_ReadMe.txt
Label1.ForeColor = oldC 'change the link color to default color
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If LoadCursor(0, 32649) > 0 Then
'show hand cursor windows 2000/XP
SetCursor LoadCursor(0, 32649)
Else
'load hand cursor from file, windows 9x
Label1.MousePointer = 99
Set Label1.MouseIcon = LoadPicture("hand.cur")
End If
Label1.ForeColor = hvrC 'enable hover color
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This UserControl surrounds the entire Label at all times
'so when the mousepointer moves off the Lable we detect that
'the UserControl now has the Mouse over it so now we change
'the text color back to its original ForeColor.
Label1.ForeColor = oldC
End Sub

Private Sub UserControl_InitProperties()
Set Font = Ambient.Font 'set the default font
Label1.BackColor = Ambient.BackColor
UserControl.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
'Here is where we set the UserControl to surround
'the entire Label at all times so when the
'mousepointer moves off the Lable we detect that
'the UserControl now has the Mouse over it and
'stop the hover color

If Label1.AutoSize = True Then
'set Lable1 position & UserContril size to suit
'this keeps the start of the Label in the same place
    Label1.Left = 120
    Label1.Top = 120
'resize the UserControl to surrond the Label
    UserControl.Width = Label1.Width + 240
    UserControl.Height = Label1.Height + 240
    Else 'no auto size
    If UserControl.Width > 240 Then
        Label1.Left = 120
        Label1.Top = 120
        Label1.Width = UserControl.Width - 240
        Else 'prevent a negitive width for Label1
        Label1.Left = 120
        Label1.Top = 120
        UserControl.Width = 270
    End If
    If UserControl.Height > 240 Then
        Label1.Height = UserControl.Height - 240
        Else 'prevent a negitive height for Label1
        UserControl.Height = 270
    Label1.Height = UserControl.Height - 240
    End If
End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'The propertybag, the essence of nearly all good UserConrols
With PropBag
    Label1.Caption = .ReadProperty("Caption", "Please click here to vote") ' reads and sets the caption to the Label1.Caption
    Label1.BackColor = .ReadProperty("BackColor", Ambient.BackColor)  'reads and sets the BackColor to the Label1.BackColor
    Label1.ForeColor = .ReadProperty("ForeColor", vbButtonText) 'reads and sets....oh you know the rest
    Label1.AutoSize = .ReadProperty("AutoSize", True)
    Label1.BorderStyle = .ReadProperty("BorderStyle", False)
    ThankyouMessage = .ReadProperty("mess", "Thankyou")
    hvrC = .ReadProperty("HoverColor", Ambient.ForeColor)
    Set Font = .ReadProperty("Font", Ambient.Font)
oldC = Label1.ForeColor 'record original link color
UserControl.BackColor = Label1.BackColor 'keeps Label & UserControl Background colors the same
End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' Write to your propertybag
With PropBag
    .WriteProperty "Caption", Label1.Caption, "Please click here to vote"     ' Write caption property
    .WriteProperty "BackColor", Label1.BackColor, Ambient.BackColor ' Write Background color property
    .WriteProperty "ForeColor", Label1.ForeColor, vbButtonText ' and so on...
    .WriteProperty "AutoSize", Label1.AutoSize, True
    .WriteProperty "BorderStyle", Label1.BorderStyle, False
    .WriteProperty "mess", ThankyouMessage, "Thankyou"
    .WriteProperty "Font", newFont, Ambient.Font
    .WriteProperty "HoverColor", hvrC, Ambient.ForeColor
Label1.BackColor = UserControl.BackColor 'keeps Label & UserControl Background colors the same
End With
End Sub
Public Property Get Caption() As Variant 'this reads from the caption property
Caption = Label1.Caption ' set the Label1.Caption to the propoperty caption
End Property
Public Property Let Caption(ByVal vNewValue As Variant) ' Set the Caption property
Label1.Caption = vNewValue ' Set's Label1.Caption to new value
PropertyChanged ' Triggers the WriteProperties event
UserControl_Resize 'resize Label1 & UserControl
End Property
Public Property Get AutoSize() As Boolean 'this reads from the AutoSize property
AutoSize = Label1.AutoSize 'Set's AutoSize to True or False
End Property
Public Property Let AutoSize(ByVal vNewValue As Boolean) 'Set the AutoSize property
Label1.AutoSize = vNewValue 'Set the AutoSize property to True or False
PropertyChanged ' Triggers the WriteProperties event
UserControl_Resize 'resize Label1 & UserControl
End Property
Public Property Get Boarder() As Boolean 'this reads from the Boarder property
Boarder = Label1.BorderStyle 'Set's Boarder to True or False
End Property
Public Property Let Boarder(ByVal vNewValue As Boolean) 'Set the Boarder property
                         'BoolToInt Returns "0" or "1" from a True or False
Label1.BorderStyle = BoolToInt(vNewValue) 'Set the Boarder property to True or False
PropertyChanged ' Triggers the WriteProperties event
UserControl_Resize
End Property
Public Property Get HoverColor() As OLE_COLOR 'this reads from the HoverColor property
HoverColor = hvrC 'Set the HoverColor
End Property
Public Property Let HoverColor(ByVal vNewValue As OLE_COLOR) 'Set the HoverColor property
hvrC = vNewValue 'Set the HoverColor property to vNewValue
PropertyChanged ' Triggers the WriteProperties event
End Property
Public Property Get BackColor() As OLE_COLOR 'this reads from the BackColor property
BackColor = Label1.BackColor 'Set the BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR) 'Set the BackColor
Label1.BackColor = vNewValue 'Set the BackColor property to vNewValue
UserControl.BackColor = vNewValue 'Set the UserControl.BackColor property to vNewValue
PropertyChanged ' Triggers the WriteProperties event
End Property
Public Property Get ForeColor() As OLE_COLOR 'this reads from the ForeColor property
ForeColor = Label1.ForeColor 'Set the ForeColor property
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR) 'Set the ForeColor
Label1.ForeColor = vNewValue 'Set the ForeColor to vNewValue
PropertyChanged ' Triggers the WriteProperties event
End Property
Public Property Get Font() As Font 'this reads from the Font property
Set Font = newFont 'Set the Font property
End Property
Public Property Set Font(ByVal fontData As Font) 'Set the Font property
Set newFont = fontData
Set Label1.Font = newFont 'Set the Font to fontData
PropertyChanged "Font" ' Triggers the WriteProperties event
UserControl_Resize 'resize Label1 & UserControl
End Property
Private Function BoolToInt(ByVal bool As Boolean) As Long
BoolToInt = IIf(bool = True, 1, 0) 'Returns "0" or "1" from a true or false
End Function
'********************************************************************************
'This next funcion is not used. Its just an
'example of the oppisite of the BoolToInt function,
'you never know it might come in handy one day.
'
'Private Function IntToBool(ByVal myInt As Long) As Boolean
'IntToBool = IIf(myInt = 1, True, False) 'Returns true or false from a "0" or "1"
'End Function
'********************************************************************************
