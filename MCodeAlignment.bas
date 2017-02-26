Attribute VB_Name = "MCodeAlignment"
Option Explicit

Sub CodeAlignment()
    Dim arr_key(1 To 30, 1 To 2) As String
    Dim dic As Object
    Dim str_code As String, arr_code
    Dim start_row As Long, end_row As Long
    Dim code_start_row As Long
    
    Set dic = CreateObject("Scripting.Dictionary")
    AddKey arr_key, dic
    
    str_code = GetSelectCode(start_row, end_row)
    Debug.Print str_code
    
    arr_code = VBA.Split(str_code, vbNewLine)
    
    code_start_row = CheckCode(arr_code)
    If code_start_row = -1 Then Exit Sub
    
    
    If OperateCode(dic, arr_code, code_start_row) Then '�������
        str_code = VBA.Join(arr_code, vbNewLine)
'        Debug.Print str_code
        
        DeleteCode start_row, end_row
        insert_code str_code
    End If
    
End Sub

'�ҵ����̻��ߺ����Ŀ�ʼλ��
Function CheckCode(arr_code) As Long
    Dim flag As Boolean
    Dim i As Long
    
    CheckCode = -1
    For i = 0 To UBound(arr_code)
        arr_code(i) = VBA.Trim(arr_code(i)) 'ȥ���ո�
        
        If arr_code(i) Like "Sub *" Then
            flag = True
        ElseIf arr_code(i) Like "Function *" Then
            flag = True
        ElseIf arr_code(i) Like "Private Sub *" Then
            flag = True
        ElseIf arr_code(i) Like "Private Function *" Then
            flag = True
        ElseIf arr_code(i) Like "Public Sub *" Then
            flag = True
        ElseIf arr_code(i) Like "Public Function *" Then
            flag = True
        End If
        
        If flag Then
            CheckCode = i
            Exit Function
        End If
        
    Next i
    
End Function

Function OperateCode(dic As Object, arr_code, code_start_row As Long) As Boolean
    Dim n_tab As Long
    Dim i As Long
    Dim i_space As Long '��1���ո���ֵ�λ��
    Dim str_key As String
    Dim tmp_tab As Long
    
    n_tab = 1
    For i = code_start_row + 1 To UBound(arr_code)
        arr_code(i) = VBA.Trim(arr_code(i)) 'ȥ���ո�
        
        If dic.Exists(arr_code(i)) Then
            '�ؼ��ֵĽ���������else������do
            If arr_code(i) = "Else" Then
                GoTo key_mid
            ElseIf arr_code(i) = "Do" Then
                GoTo key_start
            Else
                GoTo key_end
            End If
        Else
            i_space = VBA.InStr(arr_code(i), " ")
            str_key = VBA.Mid(arr_code(i), 1, i_space)
            
            If i_space > 0 Then
                If dic.Exists(str_key) Then
                    '���ڵ�keyֵ
                    tmp_tab = dic(str_key)
                    If tmp_tab = 1 Then
                        If arr_code(i) Like "If * Then *" Then  '���ܴ���ע�͵����
                            If IsAllSpace(VBA.CStr(arr_code(i))) Then
                                GoTo key_start
                            Else
                                GoTo key_not_exists
                            End If
                        Else
key_start:
                            arr_code(i) = VBA.String(n_tab, vbTab) & arr_code(i)
                            n_tab = n_tab + 1
                        End If
                    ElseIf tmp_tab = -1 Then
key_end:
                        n_tab = n_tab - 1
                        arr_code(i) = VBA.String(n_tab, vbTab) & arr_code(i)
                    ElseIf tmp_tab = 0 Then
key_mid:
                        arr_code(i) = VBA.String(n_tab - 1, vbTab) & arr_code(i)
                    End If
                Else
                    '�����ڵ�key
key_not_exists:
                    arr_code(i) = VBA.String(n_tab, vbTab) & arr_code(i)
                End If
            Else
                'û�пո�ĵ������������ע�ͻ���������stop�ȣ�
                'else��next�ȹؼ�����ǰ���Ѿ��жϹ���
                arr_code(i) = VBA.String(n_tab, vbTab) & arr_code(i)
            End If
        End If
    Next i
    
    OperateCode = True
End Function

Function AddKey(arr_key() As String, dic As Object) As Long
    Dim k As Long, i As Long
   
    k = 1
    arr_key(k, 1) = "Sub ": arr_key(k, 2) = "End Sub": k = k + 1
    arr_key(k, 1) = "Function ": arr_key(k, 2) = "End Function": k = k + 1
    arr_key(k, 1) = "If ": arr_key(k, 2) = "End If": k = k + 1
    arr_key(k, 1) = "With ": arr_key(k, 2) = "End With": k = k + 1
    arr_key(k, 1) = "For ": arr_key(k, 2) = "Next": k = k + 1          '������next�󲻼ӱ�����
    arr_key(k, 1) = "For Each ": arr_key(k, 2) = "Next ": k = k + 1     '������next��ӱ�����
    arr_key(k, 1) = "Do ": arr_key(k, 2) = "Loop ": k = k + 1             '��ʱ����ע��
    arr_key(k, 1) = "Do": arr_key(k, 2) = "Loop": k = k + 1
    arr_key(k, 1) = "While ": arr_key(k, 2) = "Wend": k = k + 1
    arr_key(k, 1) = "Select ": arr_key(k, 2) = "End Select": k = k + 1
    arr_key(k, 1) = "": arr_key(k, 2) = "Else": k = k + 1
    arr_key(k, 1) = "": arr_key(k, 2) = "Else ": k = k + 1  '��ʱ����ע��
    arr_key(k, 1) = "": arr_key(k, 2) = "ElseIf ": k = k + 1
    arr_key(k, 1) = "": arr_key(k, 2) = "Case ": k = k + 1
    arr_key(k, 1) = "": arr_key(k, 2) = "Case Else": k = k + 1
    
   '�ؼ��ֶ�Ӧ��item������Ҫ���ӵ�tab��
    For i = 1 To k - 1
    If VBA.Len(arr_key(i, 1)) Then
        dic(arr_key(i, 1)) = 1  '�ؼ��ֿ�ʼ
        dic(arr_key(i, 2)) = -1  '�ؼ��ֽ���
    Else
        dic(arr_key(i, 2)) = 0  '�ؼ����м�
    End If
    Next i
   
    AddKey = k
End Function

'�ж�If * Then *��then�����Ƿ�ȫ�ǿո���ע�ͷ���֮ǰ
Function IsAllSpace(str As String) As Boolean
    Dim i As Long
    Dim i_start As Long
    Dim str_tmp As String
    
    i_start = VBA.InStr(str, "Then ") + VBA.Len("Then ")
    
    IsAllSpace = True
    
    For i = i_start To VBA.Len(str)
        str_tmp = VBA.Mid$(str, i, 1)
        If str_tmp = "'" Then Exit For
        
        If str_tmp <> " " Then
            IsAllSpace = False
            Exit For
        End If
    Next
    
End Function

Function GetSelectCode(start_row As Long, end_row As Long) As String
    Application.VBE.ActiveCodePane.GetSelection start_row, 0, end_row, 0
    GetSelectCode = Application.VBE.ActiveCodePane.CodeModule.Lines(start_row, end_row - start_row + 1)
End Function

Function DeleteCode(start_row As Long, end_row As Long)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines start_row, end_row - start_row + 1
End Function

