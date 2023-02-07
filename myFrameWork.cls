VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myFrameWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ShtDb As Worksheet
Private KlStar As String
Private LControl As Collection

Public Function Getbaris(sht As Worksheet, KL As String) As Long
    Getbaris = sht.Cells(sht.Rows.Count, KL).End(xlUp).Row
End Function

Public Property Let SetControl(ByRef vNewValue As Control)
    LControl.Add vNewValue
End Property

Public Sub Update()
Dim Baris As Long
Dim isi() As Variant

ReDim isi(LControl.Count - 1)
Baris = Getbaris(ShtDb, KlStar) + 1
For i = 0 To LControl.Count - 1
    isi(i) = LControl(i + 1).Text
Next
ShtDb.Range(KlStar & Baris).Resize(1, UBound(isi) + 1).Value = isi
End Sub

Private Sub Class_Initialize()
    Set LControl = New Collection
End Sub

Public Sub Reset()
For Each ctrl In LControl
    ctrl.Text = vbNullString
Next
End Sub

Public Sub SetDatabase(sht As Worksheet, KL As String)
Set ShtDb = sht
KlStar = KL
End Sub
