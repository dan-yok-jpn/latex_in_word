VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public TEX2IMG As String

Private Sub Document_Open()

    Dim prop As Object
    Dim file_path As String
    Dim update As Boolean
    
    For Each prop In Me.CustomDocumentProperties
        If prop.Name = "tex2img" Then
            file_path = prop.Value
            If Not Dir(file_path) = "" Then
                Me.TEX2IMG = file_path
                Exit Sub
            Else
                update = True 'maybe moved to another directory
                Exit For
            End If
        End If
    Next

    With Application.FileDialog(msoFileDialogFilePicker)
        With .Filters
            .Clear
            .Add "excutable", "*.exe"
        End With
        If .Show = -1 Then
            file_path = .SelectedItems(1)
            .Filters.Clear
        Else
            MsgBox "Install TeX2imgc first, and try again."
            .Filters.Clear
            Exit Sub
        End If
    End With

    If InStr(file_path, "TeX2imgc") > 0 Then
        If update Then
            Me.CustomDocumentProperties("tex2img").Value = file_path
        Else
            With Me.CustomDocumentProperties
                .Add _
                    Name:="tex2img", _
                    LinkToContent:=False, _
                    Type:=msoPropertyTypeString, _
                    Value:=file_path
            End With
        End If
        Me.TEX2IMG = file_path
        Me.Save
    End If

End Sub


