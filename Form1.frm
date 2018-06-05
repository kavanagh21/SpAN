VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Processing"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   " Data "
      Height          =   3135
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton Command8 
         Caption         =   "Subsample based on list"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Find"
         Height          =   315
         Left            =   240
         TabIndex        =   39
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   240
         TabIndex        =   37
         Text            =   "Not set"
         Top             =   1680
         Width           =   5055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Copy data to clipboard"
         Height          =   255
         Left            =   7200
         TabIndex        =   30
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find"
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "Not set"
         Top             =   720
         Width           =   5055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   5520
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   7200
         TabIndex        =   15
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Copy graph to clipboard"
         Height          =   495
         Left            =   4560
         TabIndex        =   41
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Video file:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Data file:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "-"
         Height          =   255
         Left            =   5520
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Average flux:"
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "-"
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Points:"
         Height          =   255
         Left            =   5520
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Show"
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Data points in memory: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display Settings "
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3375
      Begin VB.CheckBox Check7 
         Caption         =   "Check6"
         Height          =   255
         Left            =   2160
         TabIndex        =   46
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check6"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   2040
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0026
         Left            =   1920
         List            =   "Form1.frx":0039
         TabIndex        =   44
         Text            =   "14"
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Draw point marker"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2520
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   50
         Left            =   1800
         Max             =   1000
         Min             =   1
         SmallChange     =   5
         TabIndex        =   34
         Top             =   960
         Value           =   1
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Text            =   "25"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   31
         Text            =   "100"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Draw slice number"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Highlight inversions"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Y grid"
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show X grid"
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   "5"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "5000"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Grid font size"
         Height          =   255
         Left            =   1920
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Grid X:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Grid Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "XStep"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Hi"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Low"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   13935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   13875
      TabIndex        =   1
      Top             =   3600
      Width           =   13935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   315
      Left            =   5160
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Send form to printer"
      Height          =   495
      Left            =   360
      TabIndex        =   42
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox prtlist 
      Height          =   315
      Left            =   4080
      TabIndex        =   36
      Text            =   "Combo2"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   4
      Top             =   10200
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   10200
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flux(100000) As Double
Dim fluxid(100000) As String
Dim dpim As Integer




Private Sub Check1_Click()
DrawGraph

End Sub

Private Sub Check2_Click()
DrawGraph
End Sub

Private Sub Check3_Click()
DrawGraph

End Sub

Private Sub Check4_Click()
DrawGraph

End Sub

Private Sub Check5_Click()
DrawGraph

End Sub

Private Sub Combo1_Click()
List1.Clear

If Combo1.List(Combo1.ListIndex) = "Down points" Then
For n = 1 To dpim
    If fluxid(n) = "up->down" Then List1.AddItem n & ": " & flux(n): ttl = ttl + flux(n): ttc = ttc + 1
Next n
Else
For n = 1 To dpim
    If fluxid(n) = "down->up" Then List1.AddItem n & ": " & flux(n): ttl = ttl + flux(n): ttc = ttc + 1
Next n
End If
If ttc = 0 Then Exit Sub

Label10.Caption = ttc
Label12.Caption = Round(ttl / ttc, 1)

End Sub

Private Sub Command1_Click()
Dim cln() As String
If Dir(Text1.Text) = "" Then Exit Sub
If Text1.Text = "" Then Exit Sub

Open Text1.Text For Input As #1

Line Input #1, waste
Line Input #1, waste
Line Input #1, waste
Line Input #1, waste

Do While Not EOF(1)
    Line Input #1, ltx
    cln = Split(ltx, vbTab)
    pos = pos + 1
    flux(pos) = Val(cln(2))
Loop

Close #1

Label6.Caption = "Data points in memory: " & pos - 1
dpim = pos - 1
Text4.Text = Int(Picture1.Width / dpim)
HScroll2.Value = Int(Picture1.Width / dpim)
 


End Sub

Private Sub Command2_Click()

Text1.Text = OpenFileAPI("", "", "", False, "")
Command1_Click
Command3_Click
Combo1.ListIndex = 1
Combo1_Click


End Sub

Private Sub DrawGraph()

Picture1.Cls

'work out the step distance so the scroll bar works; fpr = the number of points per width of screen
fpr = Int(Picture1.Width / Val(Text4.Text))
'udp = undisplayable data points, e.g. the number of points in memory minus those displayed on the screen
udp = dpim - fpr
If udp > 0 Then
    HScroll1.Min = 0
    HScroll1.Max = udp
    HScroll1.SmallChange = Int(udp / 500) + 1
    HScroll1.LargeChange = Int(udp / 10) + 1
    HScroll1.Enabled = True
    startpt = 1 + (HScroll1.Value - 1)
Else
    HScroll1.Enabled = False
    startpt = 1
End If



Label7.Caption = 1 + startpt
lpt = -1

If Check1.Value = 1 Then
    For n = 0 To fpr Step Val(Text6.Text)
        newX = Int(Val(Text4.Text) * n)
        Picture1.Line (newX, 0)-(newX, Picture1.Height), &H404040
    Next n
End If

If Check2.Value = 1 Then
    For n = Val(Text2.Text) To Val(Text3.Text) Step Val(Text5.Text)
        'work out where this line would be
        'these are FLUX values
        picbr = Picture1.Height
        limrng = Val(Text3.Text) - Val(Text2.Text)
        calcflux = n - Val(Text2.Text)
        calcflux = calcflux * (picbr / limrng)
        fluxY = picbr - calcflux
        Picture1.Line (0, fluxY)-(Picture1.Width, fluxY), &H404040
        Picture1.CurrentX = 0
        Picture1.CurrentY = fluxY
        Picture1.ForeColor = vbWhite
        Picture1.FontSize = Combo2
        Picture1.Print n & " ^"
    Next n
End If


For n = 1 + startpt To dpim
    xps = xps + 1
    newX = Int(Val(Text4.Text) * xps)
    If newX > Picture1.Width And lpt = -1 Then lpt = n
    
    
    'work out where in the range this falls.
    picbr = Picture1.Height
    limrng = Val(Text3.Text) - Val(Text2.Text)
    calcflux = flux(n) - Val(Text2.Text)
    calcflux = calcflux * (picbr / limrng)
    newy = picbr - calcflux
    
    If fpt = 0 Then lastY = newy: fpt = 1
    
    'each position needs to be defined as either moving upward, or downward.
    If flux(n) > prevFlux Then direction = "up": lcol = vbYellow
    If flux(n) < prevFlux Then direction = "down": lcol = vbRed
    fluxid(n) = "normal"
    
    If prevDirection = "down" And direction = "up" And n > 3 Then
        'this point is an inversion in the plot
        'some of these are not valid though and should be rejected.
        'the next two points should both in the same trajectory for this to count.
        'in the case of a down to up switch, the next two points should be upwards of this.
        frp = n - 1
        If flux(frp) < flux(frp + 1) And flux(frp + 1) < flux(frp + 2) And flux(frp) < flux(frp - 1) And flux(frp - 1) < flux(frp - 2) Then
            'the next point is higher than this, and the point after is higher than that. This appears a real directional change.
            ccol = vbBlue
            fluxid(frp) = "down->up"
        Else
            'the trajectory doesn't look right
            ccol = &H8000000F
        End If
        
        
        If Check3.Value = 1 And Check6.Value = 0 Then Picture1.Circle (lastX, lastY), 100, ccol
    
    End If
    
    If prevDirection = "up" And direction = "down" And n > 3 Then
        'this point is an inversion in the plot
        'some of these are not valid though and should be rejected.
        'we're now going down, so the next two points should be lower.
        frp = n - 1
        If flux(frp) > flux(frp + 1) And flux(frp + 1) > flux(frp + 2) And flux(frp) > flux(frp - 1) And flux(frp - 1) > flux(frp - 2) Then
            'the next point is higher than this, and the point after is higher than that. This appears a real directional change.
            ccol = vbGreen
            fluxid(frp) = "up->down"
        Else
            'the trajectory doesn't look right
            ccol = &H8000000F
        End If
        
        
        
        If Check3.Value = 1 And Check7.Value = 0 Then Picture1.Circle (lastX, lastY), 100, ccol
    
    
    End If






    Picture1.Line (lastX, lastY)-(newX, newy), lcol
    If Check5.Value = 1 And (Check7.Value = 0 And Check6.Value = 0) Then Picture1.Circle (newX, newy), 20, vbWhite
    If Check4.Value = 1 Then
        Picture1.CurrentX = newX
        Picture1.CurrentY = newy
        Picture1.ForeColor = vbWhite
        Picture1.FontName = "Arial"
        Picture1.FontSize = 8
        Picture1.Print n
    End If
    
        
    
    lastX = newX
    lastY = newy
    prevFlux = flux(n)
    prevDirection = direction
Next
If lpt = -1 Then lpt = dpim
Label8.Caption = lpt


End Sub

Private Sub Command3_Click()
DrawGraph

End Sub

Private Sub Command4_Click()
Dim stc As String
For n = 0 To List1.ListCount - 1
flid = Val(List1.List(n))
stc = stc & flid & vbTab & flux(flid) & vbTab & fluxid(flid) & vbCrLf
Next n
Clipboard.Clear
Clipboard.SetText CStr(stc)

End Sub

Private Sub Command5_Click()
Clipboard.Clear
Clipboard.SetData Picture1.Picture, vbCFBitmap

End Sub

Private Sub Command6_Click()
    For Each prt In Printers
        If prt.DeviceName = prtlist.List(prtlist.ListIndex) Then
            Set Printer = prt
            Exit For
        End If
    Next prt
    
    Printer.ColorMode = 2
    
    
    
    Me.PrintForm

End Sub

Private Sub Command7_Click()
Text7.Text = OpenFileAPI("", "", "", False, "")
End Sub

Private Sub Command8_Click()
Open Environ("TEMP") & "\span\dovid.bat" For Output As #1
    Print #1, "mkdir %TEMP%\span"
    Print #1, "del /Q %TEMP%\span\*.jpg"
    Print #1, Chr(34) & App.Path & "\ffmpeg" & Chr(34) & " -i " & Chr(34) & Text7.Text & Chr(34) & " %temp%\span\%%d.jpg"
    
    For n = 0 To List1.ListCount - 1
        Print #1, "rename %TEMP%\span\" & Val(List1.List(n)) & ".jpg k" & n + 1 & ".jpg"
        fullstr = fullstr & "-i %TEMP%\span\" & Val(List1.List(n)) & ".jpg "
    Next n
    Print #1, Chr(34) & App.Path & "\ffmpeg" & Chr(34) & " -r 10 -i %temp%\span\k%%d.jpg -c:v libx264 -vf fps=25 -pix_fmt yuv420p %TEMP%\span\out.mp4 -y"
    Print #1, "copy " & Chr(34) & "%temp%\span\out.mp4" & Chr(34) & " " & Chr(34) & "%userprofile%\desktop\SpAn Video.mp4" & Chr(34) & ""
    Print #1, "pause"
    
Close #1
res = Shell(Environ("TEMP") & "\span\dovid.bat", vbNormalFocus)

End Sub

Private Sub Form_Load()
Dim prt As Printer

    For Each prt In Printers
        prtlist.AddItem prt.DeviceName
    Next prt
    
    prtlist.ListIndex = 0
    
Text2.Text = GetSetting("dta", "settings", "low")
Text3.Text = GetSetting("dta", "settings", "high")
Text4.Text = GetSetting("dta", "settings", "xstep")

End Sub

Private Sub Form_Resize()
On Error Resume Next

Picture1.Width = Form1.Width - 500
Picture1.Height = Form1.Height - Picture1.Top - 1200
Label8.Left = Form1.Width - Label8.Width - 500
Label8.Top = Form1.Height - Label8.Height - 650
Label7.Top = Form1.Height - Label7.Height - 650

End Sub

Private Sub HScroll1_Change()
DrawGraph

End Sub

Private Sub HScroll2_Change()
Text4.Text = HScroll2.Value
DrawGraph

End Sub

Private Sub R_Click()
Command3_Click

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label7_DblClick()
HScroll1.Value = InputBox("Enter new val", , HScroll1.Value)

End Sub

Private Sub Text2_Change()
SaveSetting "dta", "settings", "low", Text2.Text

End Sub

Private Sub Text3_Change()
SaveSetting "dta", "settings", "high", Text3.Text

End Sub

Private Sub Text4_Change()
SaveSetting "dta", "settings", "xstep", Text4.Text

End Sub

