VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IE_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"

Private WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Private bDownloading As Boolean

Public Property Set IEctl(IncomeIE As SHDocVw.InternetExplorer)
   Set IE = IncomeIE
End Property

Private Sub Class_Terminate()
'  On Error Resume Next
  Set IE = Nothing
End Sub

Private Sub IE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
   Dim s As String
   If Not bDownloading Then
      s = Time & vbTab & "Navigation begin."
      s = s & vbCrLf & vbTab & "URL = " & CStr(URL)
      AddToLog s
   End If
   bDownloading = True
End Sub

Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  Dim s As String
  If pDisp Is IE Then
     s = Time & vbTab & "Documente complete."
     s = s & vbCrLf & vbTab & "URL = " & IE.LocationURL
     Do While IE.Busy
        DoEvents
     Loop
     s = s & vbCrLf & vbTab & "Location = " & IE.LocationName
     AddToLog s
     bDownloading = False
  End If
End Sub

Private Sub IE_DownloadBegin()
   Dim s As String
   If Not bDownloading Then
      s = Time & vbTab & "Navigation begin."
      s = s & vbCrLf & vbTab & "URL = " & IE.LocationURL
      AddToLog s
   End If
   bDownloading = True
End Sub
