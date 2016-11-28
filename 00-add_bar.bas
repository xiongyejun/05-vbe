Attribute VB_Name = "模块1"
Option Explicit

'引用 microsoft visual basic for application extensibility 5.3

Dim only As New Collection


Sub main()
    If Not check_vbproject Then
        Exit Sub
    End If
    
    add_bar
End Sub

Function add_bar()
    Dim bar_btn As CommandBarButton
    Dim mybar_event As mybar
    Dim str As String, arr_sr
    Dim i As Long
    Const VBE_DIR As String = "\04-github\05-vbe\vbe_code\CommandBarButton"
        
    On Error Resume Next
    Application.VBE.CommandBars("mybar").Delete
    On Error GoTo 0
    
    str = fso_read_txt(get_my_doc() & VBE_DIR)
    arr_sr = Split(str, vbNewLine)
    
    With Application.VBE.CommandBars.Add
        .NameLocal = "mybar"
        .Visible = True
'        .Position = msoBarTop
        .Left = 1100
        .Top = 400
        
        For i = 0 To UBound(arr_sr)
            If arr_sr(i) <> "" Then
                Set mybar_event = New mybar
                Set bar_btn = .Controls.Add
                bar_btn.Caption = arr_sr(i)
                bar_btn.Style = msoButtonCaption
                bar_btn.BeginGroup = True
                Set mybar_event.myevent = Application.VBE.Events.CommandBarEvents(bar_btn)
                only.Add mybar_event
            End If
        Next i
        
        .Width = 50 '小的让他自动调整
    End With
    
    

    Set bar_btn = Nothing
    Set mybar_event = Nothing
'
End Function

Function fso_read_txt(file_name As String) As String
    Dim fso As Object, sr As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sr = fso.OpenTextFile(file_name, 1) 'ForReading=1

    fso_read_txt = sr.ReadAll()
    
    Set fso = Nothing
    Set sr = Nothing
End Function

Function insert_code(str_code As String)
    Dim i_row As Long
    
    Application.VBE.ActiveCodePane.GetSelection i_row, 0, 0, 0
    Application.VBE.SelectedVBComponent.CodeModule.InsertLines i_row, str_code
End Function

Function check_vbproject() As Boolean
    Dim obj As Object
    
    On Error Resume Next
    Set obj = Application.VBE.ActiveVBProject
    If Err.Number <> 0 Then
        MsgBox "请勾选 信任对VBA工程对象模型的访问"
        check_vbproject = False
    Else
        check_vbproject = True
    End If
End Function

Function get_my_doc() As String
    Dim wsh As Object
    Dim str As String
    
    Set wsh = CreateObject("WScript.Shell")
    str = wsh.SpecialFolders("Mydocuments")
    get_my_doc = str
    
    Set wsh = Nothing
End Function


