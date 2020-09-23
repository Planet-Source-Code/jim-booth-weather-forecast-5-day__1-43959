VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "5 Day Weather Forecast"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox chkSaveSetting 
      BackColor       =   &H80000013&
      Caption         =   "Remember my settings"
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   2295
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
      Begin VB.OptionButton optFaren 
         BackColor       =   &H80000013&
         Caption         =   "Fahrenheit"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optCel 
         BackColor       =   &H80000013&
         Caption         =   "Celsius"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   855
      End
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6600
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblPlace 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Image imgButton 
      Height          =   510
      Left            =   120
      Picture         =   "Form1.frx":0ECA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Surl As String
Dim sTemp As String
Dim sTemp2  As String

Dim lctr As Long
Dim lctr2 As Long
Dim lTime As Long
Dim bDone As Boolean

Private WithEvents IEDoc As MSHTML.HTMLDocument
Attribute IEDoc.VB_VarHelpID = -1

Private Sub cmdClose_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    
    bDone = False
    lctr = 0
    lctr2 = 0
    lTime = Timer
    Inet1.Cancel
    Inet1.RequestTimeout = 30 '30 second timeout
    Inet1.AccessType = icUseDefault
    Screen.MousePointer = vbHourglass
    lblPlace = "Querying web server..."
    sTemp = Inet1.OpenURL("http://weather.yahoo.com/search/weather2?p=" & txtSearch, icString)

    If sTemp = "" Then
        lblPlace = "No response from server"
        Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    If InStr(sTemp, "location matches") Then
        Screen.MousePointer = vbDefault
        lblPlace = "Please select your area"
    Else
        Open "c:\temp.html" For Output As #1
        sTemp = Replace(sTemp, Chr(10), vbCrLf)
        Print #1, sTemp
        Close #1
        
        Open "c:\temp.html" For Input As #1
        
        Do While Not EOF(1)
            Line Input #1, sTemp2
            
            If InStr(LCase(sTemp2), "http://mtf.news.yahoo.com/mailto?") Then
                lctr = InStr(sTemp2, "url") + 4
                lctr2 = InStr(lctr + 4, sTemp2, ".htm")
                Surl = Mid(sTemp2, lctr, lctr2 - lctr + 5)
            End If
            
            If InStr(LCase(sTemp2), "<!--forecast header-->") Then
                Line Input #1, sTemp2
                Line Input #1, sTemp2
                Line Input #1, sTemp2
                
                lctr = InStr(sTemp2, "<b>")
                lctr2 = InStr(sTemp2, "</b>")
                lblPlace = Mid(sTemp2, lctr + 3, lctr2 - lctr - 3)
                If chkSaveSetting.Value = True Then
                    SaveSetting App.Title, "Settings", "Place", lblPlace
                    SaveSetting App.Title, "Settings", "Save", "True"
                Else
                    SaveSetting App.Title, "Settings", "Save", "True"
                End If
                
            End If
            
        Loop
                
        Close #1
    End If
    
    Screen.MousePointer = vbDefault
    
    Call GetMatchesOrWeather(sTemp)

End Sub

Private Sub Form_Load()
    On Error Resume Next
    If GetSetting(App.Title, "Settings", "Save", "False") = "True" Then
        chkSaveSetting.Value = vbChecked
        txtSearch = GetSetting(App.Title, "Settings", "Place", "")
    End If
End Sub

Private Function IEDoc_onclick() As Boolean
    On Error Resume Next
    If Not bDone Then
        lblPlace = "Querying server for : " & IEDoc.activeElement.innerText
        bDone = True
    End If
    
    If IEDoc.activeElement.getAttribute("Href") <> "" Then
        Screen.MousePointer = vbHourglass
        Web.Navigate2 IEDoc.activeElement.getAttribute("Href")
        If chkSaveSetting.Value = vbChecked Then
            SaveSetting App.Title, "Settings", "Place", IEDoc.activeElement.innerText
            SaveSetting App.Title, "Settings", "Save", "True"
        Else
            SaveSetting App.Title, "Settings", "Save", "False"
        End If
    End If
    
End Function

Private Function IEDoc_oncontextmenu() As Boolean
    On Error Resume Next
    IEDoc_oncontextmenu = False
End Function

Private Function IEDoc_onstop() As Boolean
    On Error Resume Next
    Screen.MousePointer = vbDefault
End Function

Private Sub optCel_Click()
    On Error Resume Next
    If Surl <> "" Then Call GetGfxForcast(Surl)
End Sub

Private Sub optFaren_Click()
    On Error Resume Next
    If Surl <> "" Then Call GetGfxForcast(Surl)
End Sub

Private Sub txtSearch_GotFocus()
    On Error Resume Next
    cmdSearch.Default = True
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch)
End Sub

Private Sub Web_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    On Error Resume Next
    If Not LCase(URL) = "c:\temp.html" Then
        Surl = URL
        Cancel = True
        Call GetGfxForcast(Surl)
    End If
End Sub

'Private Sub GetTextForcast(Surl As String)
'
'    'Surl = Replace(Surl, "_f.html", "_c.html#text")
'    If optCel.Value = True Then
'        Surl = Replace(Surl, "_f.html", "_c.html#force_units=1")
'    Else
'        Surl = Replace(Surl, "_c.html", "_f.html#force_units=1")
'    End If
'
'    Inet1.AccessType = icUseDefault
'    sTemp = Inet1.OpenURL(Surl, icString)
'
'    lctr = 0
'    lctr2 = 0
'    lctr = InStr(sTemp, "<!--TEXT FORECAST-->")
'
'    If lctr <> 0 Then
'        lctr2 = InStr(lctr, sTemp, "<!--ENDTEXT FORECAST-->")
'        If lctr2 <> 0 Then
'            sTemp = Mid(sTemp, lctr, lctr2 - lctr)
'            sTemp = Replace(sTemp, Chr(10), vbCrLf)
'            Open "c:\temp.html" For Output As #1
'            Print #1, sTemp
'            Close #1
'            Web.Navigate2 "c:\temp.html"
'        End If
'    End If
'
'End Sub

Private Sub GetGfxForcast(Surl As String)
    On Error Resume Next
    
    Screen.MousePointer = vbDefault
    
    If LCase(Surl) = "http:///" Then Exit Sub
    If optCel.Value = True Then
        Surl = Replace(Surl, "_f.html", "_c.html#force_units=1")
    Else
        Surl = Replace(Surl, "_c.html", "_f.html#force_units=1")
    End If
    If chkSaveSetting.Value = vbChecked Then
        SaveSetting App.Title, "Settings", "Link", Surl
        SaveSetting App.Title, "Settings", "Save", "True"
    Else
        SaveSetting App.Title, "Settings", "Save", "False"
    End If
    
    Inet1.AccessType = icUseDefault
    sTemp = Inet1.OpenURL(Surl, icString)
    
    lctr = 0
    lctr2 = 0
    lctr = InStr(sTemp, "<!----------------------- FORECAST ------------------------->")
    
    If lctr <> 0 Then
        lctr2 = InStr(lctr, sTemp, "<!--ENDFC-->")
        If lctr2 <> 0 Then
            sTemp = Mid(sTemp, lctr, lctr2 - lctr)
            sTemp = Replace(sTemp, Chr(10), vbCrLf)
            
            lctr = InStr(sTemp, "6-10")
            If lctr <> 0 Then
                lctr = InStrRev(LCase(sTemp), "<td ", lctr)
                If lctr <> 0 Then
                    lctr2 = InStr(lctr, LCase(sTemp), "</td>")
                    If lctr2 <> 0 Then
                        sTemp2 = Mid(sTemp, lctr, lctr2 - lctr + 5)
                        sTemp = Replace(sTemp, sTemp2, "<!-- " & sTemp2 & " -->")
                    End If
                End If
            End If
            
            lctr = InStrRev(LCase(sTemp), "<td align=center valign=middle>")
            If lctr <> 0 Then
                lctr2 = InStr(lctr, LCase(sTemp), "</table>")
                If lctr2 <> 0 Then
                    sTemp2 = Mid(sTemp, lctr, lctr2 - lctr + 8)
                    sTemp = Replace(sTemp, sTemp2, "")
                End If
            End If
            
            Open "c:\temp.html" For Output As #1
            Print #1, sTemp
            Close #1
            Web.Navigate2 "c:\temp.html"
        End If
    End If
    
End Sub

Private Sub GetMatchesOrWeather(ByVal sData As String)
    
    On Error Resume Next
    
    If InStr(sTemp, "location matches") Then
        lctr = InStr(sTemp, "location matches")
        If lctr <> 0 Then
            lctr = InStr(lctr, sTemp, "<ul>")
            If lctr <> 0 Then
                lctr2 = InStr(lctr, sTemp, "</ul>")
                sTemp = Mid(sTemp, lctr, lctr2 - lctr + 5)
                sTemp = Replace(sTemp, Chr(10), vbCrLf)
                
                Open "c:\temp.html" For Output As #1
                Print #1, sTemp
                Close #1
                
                Web.Navigate2 "c:\temp.html"
                
            End If
        Else
            
            lctr = InStr(sTemp, "<!--TEXT FORECAST-->")
            
            If lctr <> 0 Then
                lctr2 = InStr(lctr, sTemp, "<!--ENDTEXT FORECAST-->")
                If lctr2 <> 0 Then
                    sTemp = Mid(sTemp, lctr, lctr2 - lctr)
   
                            
                    sTemp = Replace(sTemp, Chr(10), vbCrLf)
                    Open "c:\temp.html" For Output As #1
                    Print #1, sTemp
                    Close #1
                    Web.Navigate2 "c:\temp.html"
                End If
            End If
        End If
    Else
        lctr = 0
        lctr2 = 0
        lctr = InStr(sTemp, "<!----------------------- FORECAST ------------------------->")
        
        If lctr <> 0 Then
            lctr2 = InStr(lctr, sTemp, "<!--ENDFC-->")
            If lctr2 <> 0 Then
                sTemp = Mid(sTemp, lctr, lctr2 - lctr)
                
                lctr = InStr(sTemp, "6-10")
                If lctr <> 0 Then
                    lctr = InStrRev(LCase(sTemp), "<td ", lctr)
                    If lctr <> 0 Then
                        lctr2 = InStr(lctr, LCase(sTemp), "</td>")
                        If lctr2 <> 0 Then
                            sTemp2 = Mid(sTemp, lctr, lctr2 - lctr + 5)
                            sTemp = Replace(sTemp, sTemp2, "<!-- " & sTemp2 & " -->")
                        End If
                    End If
                End If
                
                lctr = InStrRev(LCase(sTemp), "<td align=center valign=middle>")
                If lctr <> 0 Then
                    lctr2 = InStr(lctr, LCase(sTemp), "</table>")
                    If lctr2 <> 0 Then
                        sTemp2 = Mid(sTemp, lctr, lctr2 - lctr + 8)
                        sTemp = Replace(sTemp, sTemp2, "")
                    End If
                End If
                
                sTemp = Replace(sTemp, Chr(10), vbCrLf)
                Open "c:\temp.html" For Output As #1
                Print #1, sTemp
                Close #1
                Web.Navigate2 "c:\temp.html"
            End If
        Else
            Open "c:\temp.html" For Output As #1
            Print #1, "<center><B>No matches found.</b></center>"
            Close #1
            Web.Navigate2 "c:\temp.html"
        End If
'
'        lctr = InStr(sTemp, "<!--TEXT FORECAST-->")
'
'        If lctr <> 0 Then
'            lctr2 = InStr(lctr, sTemp, "<!--ENDTEXT FORECAST-->")
'            If lctr2 <> 0 Then
'                sTemp = Mid(sTemp, lctr, lctr2 - lctr)
'                Open "c:\temp.html" For Output As #1
'                Print #1, sTemp
'                Close #1
'                Web.Navigate2 "c:\temp.html"
'            End If
'        Else
'            'Debug.Print sTemp
'            Open "c:\temp.html" For Output As #1
'            Print #1, "<center><B>No matches found.</b></center>"
'            Close #1
'            Web.Navigate2 "c:\temp.html"
'        End If
        
    End If

End Sub

Private Sub Web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    
    Screen.MousePointer = vbDefault
    
    Set IEDoc = Nothing
    Set IEDoc = Web.Document
    
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    lblPlace = Mid(lblPlace, InStr(lblPlace, ":") + 1, Len(lblPlace))
    
End Sub
