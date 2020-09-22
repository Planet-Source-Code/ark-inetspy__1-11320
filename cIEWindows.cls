VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIEWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"IE_Class"
Attribute VB_Ext_KEY = "Member0" ,"IE_Class"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents winShell As SHDocVw.ShellWindows
Attribute winShell.VB_VarHelpID = -1
Private mCol As Collection
Dim bRefreshing As Boolean, bInit As Boolean

Private Function Add(IEctl As SHDocVw.InternetExplorer) As IE_Class
    Dim objNewMember As IE_Class
    Set objNewMember = New IE_Class
    Set objNewMember.IEctl = IEctl
    mCol.Add objNewMember
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get IE(vntIndexKey As Variant) As IE_Class
Attribute IE.VB_UserMemId = 0
  Do While bRefreshing
     DoEvents
  Loop
  Set IE = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = mCol.Count
End Property

Private Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
  Refresh_Col
End Sub

Private Sub Class_Terminate()
   Set mCol = Nothing
   Set winShell = Nothing
End Sub

Private Sub winShell_WindowRegistered(ByVal lCookie As Long)
  Refresh_Col
  Do While bRefreshing
     DoEvents
  Loop
End Sub

Private Sub winShell_WindowRevoked(ByVal lCookie As Long)
  Refresh_Col
  Do While bRefreshing
     DoEvents
  Loop
End Sub

Private Sub Refresh_Col()
   bRefreshing = True
   Dim SWs As New SHDocVw.ShellWindows
   Dim var As SHDocVw.InternetExplorer
   Set mCol = Nothing
   Set mCol = New Collection
   For Each var In SWs
       Add var
   Next
   If Not bInit Then Set winShell = SWs
   bInit = True
   Set SWs = Nothing
   Set var = Nothing
   bRefreshing = False
End Sub

