

Sub RedirectMail()
    '
    Dim FwEmail As Outlook.MailItem
    Dim EmailReadyToSend As Outlook.MailItem
    Dim t As Integer
    Dim i As Integer
    Dim j As Integer
    Dim table1 As Object
    Dim FwEmailDate As Date
    'Dim ohtml As MSHTML.MSHTML.HTMLDocument
    'Set ohtml = New MSHTML.HTMLDocument
    'Dim o_html_el As MSHTML.IHTMLElementCollection
    
    
    
    
    Dim mail_rows As Variant
    ' 此变量为处理邮件内容
    
    Dim Added_String1 As String
    Dim Added_String2 As String
    Dim p_class As String
    Dim span_style As String
    
    
    ' add some contect as fwbody(with same style)
    ' 为转发邮件内容加入一段带格式的文字
    Dim dear_all As String
    Dim row1_sentence As String
    Dim row2_sentence As String
    Dim thanks As String
    
    p_class = "MsoNormal"
    span_style = "font-family:&quot;Andalus&quot;,serif;color:#0070C0"
    dear_all = "Dear All,"
    
    
    row1_sentence = "Below is Daily approved and next day pending/approved change report related to Asia only (Yellow entry means the risk or impact of this change is high/medium.)"
    row2_sentence = "Please do review the changes impact your country and sites, and connect to your partners/ customers if necessary."
    
    
    Added_String1 = "<p class= " & Chr(34) & p_class & Chr(34) & ">" & "<span style=" & Chr(34) & span_style & Chr(34) & ">" & row1_sentence & "<o:p></o:p></span></p>"
    Added_String2 = "<p class= " & Chr(34) & p_class & Chr(34) & ">" & "<span style=" & Chr(34) & span_style & Chr(34) & ">" & row2_sentence & "<o:p></o:p></span></p>"
    dear_all = "<p class= " & Chr(34) & p_class & Chr(34) & ">" & "<span style=" & Chr(34) & span_style & Chr(34) & ">" & "Dear All," & "<o:p></o:p></span></p>"
    thanks = "<p class= " & Chr(34) & p_class & Chr(34) & ">" & "<span style=" & Chr(34) & span_style & Chr(34) & ">" & "Thanks" & "<o:p></o:p></span></p>"
    
    
    'User input
    'MsgBox ("请输入要转发邮件来自哪个邮箱（完整邮箱地址）。")
    'FwFromWho = InputBox("请输入要转发邮件来自哪个邮箱（完整邮箱地址）。")
    'FwEmailDate = Day(Date)  '默认是取自今天的邮件  可以考虑  DateTime.Now.AddDays() 更改
    'FwFromTo = InputBox("请输入此邮件转发到哪里（完整邮箱地址）。")
    'FwEmailSubject = InputBox("请输入邮件主题的关键字段，用来筛选")

    ' 在本机中搜索邮件

    Dim objInbox As Outlook.Folder

    Set objInbox = Session.GetDefaultFolder(olFolderInbox)
    
    Dim subFolder As Outlook.Folder
    
    For Each subFolder In objInbox.Folders
        'MsgBox (subFolder.Name)  ' It works
        If InStr(CStr(subFolder.Name), "****") Then '找到还含有‘****’关键字的邮件文件夹
            'MsgBox ("Catch software") ' works
            For Each FwEmail In subFolder.Items
                If DateDiff("d", FormatDateTime(FwEmail.ReceivedTime, vbShortDate), FormatDateTime(Now, vbShortDate)) < 2 Then  '距离今天 XX 天内的邮件
                    'MsgBox (FwEmail.Subject)
                    'MsgBox (FormatDateTime(Now, vbShortDate))
                    If InStr(FwEmail.Subject, "Daily Approved & Next Day Pending/Approved Change Report") Then
                    ' 在这个文件夹里，找到标题中含有'关键字'的的邮件 ----- 也可以加上 一个判定： 由谁发给我的

                        'MsgBox DateDiff("d", FormatDateTime(FwEmail.ReceivedTime, vbShortDate), FormatDateTime(Now, vbShortDate))
                        
                        '开始制作需要转发出的邮件
                        Set EmailReadyToSend = FwEmail.Forward
                        
                        '去掉不必要的表格标题
                        EmailReadyToSend.HTMLBody = Replace(EmailReadyToSend.HTMLBody, "Tower Summary:", "")
                        
                        '加入转发的邮件头&正文
                        EmailReadyToSend.HTMLBody = dear_all & Added_String1 & Added_String2 & thanks & EmailReadyToSend.HTMLBody

                        EmailReadyToSend.To = "****@*****.com"
                        EmailReadyToSend.CC = "****@*****.com"   '要cc 给谁就在这里填。
                        
                        'EmailReadyToSend.BodyFormat = FwEmail.Forward.BodyFormat
                        
                        '================处理邮件中的表格============================
                        '
                        '删除 带有 NA, EMEA, LA 记录
                        'MsgBox (EmailReadyToSend.GetInspector.WordEditor.tables.Count)  '获取邮件正文中表格的调用
                        
                        
                        'MsgBox (EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count)  ' 行数  i
                        'MsgBox (EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count) '列数  j
                        'MsgBox (EmailReadyToSend.GetInspector.WordEditor.tables(1).rows(6).Cells(8))  'cells


                        t = 1
                        Do
                        On Error Resume Next
                        ' ------处理表中的 NA EMEA LA------
                            '----NA-----
                            ' 先找到每个表头中的location  用 j 来取值 |  用 tables(t)  来迭代表格
                            '------由于循环导致调用漏洞，用On Error Resume Next 忽略报错，此项排除工作反复执行3 次
                        
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理

                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "NA") Then   '筛选出NA

                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                            
                                        End If
                                    Next

                                End If
                            Next
                            
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理

                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "NA") Then   '筛选出NA

                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                            
                                        End If
                                    Next

                                End If
                            Next

                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理

                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "NA") Then   '筛选出NA

                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                            
                                        End If
                                    Next

                                End If
                            Next
                            
                            '----/NA-----
                            '----LA-----
                            ' 先找到每个表头中的location  用 j 来取值 |  用 tables(t)  来迭代表格
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理
                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "LA") Then   '筛选出LA
                                            '删除
                                            
                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                        End If
                                    Next
                                End If
                            Next
                            
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理
                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "LA") Then   '筛选出LA
                                            '删除
                                            
                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                        End If
                                    Next
                                End If
                            Next
                            
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理
                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "LA") Then   '筛选出LA
                                            '删除
                                            
                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                        End If
                                    Next
                                End If
                            Next
                            '----/LA-----
                            
                            '----EMEA-----
                            ' 先找到每个表头中的location  用 j 来取值 |  用 tables(t)  来迭代表格
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理
                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "EMEA") Then   '筛选出EMEA
                                            '删除
                                            
                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                        End If
                                    Next
                                End If
                            Next
                            
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理
                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "EMEA") Then   '筛选出EMEA
                                            '删除
                                            
                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                        End If
                                    Next
                                End If
                            Next
                            
                            For j = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).Columns.Count  ' tables（t）每一列
                                If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(j), "LOCATION") Then '找到表row(1)的Location 再对此列中 NA EMEA，LA 处理
                                    For i = 1 To EmailReadyToSend.GetInspector.WordEditor.tables(t).rows.Count
                                        If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j), "EMEA") Then   '筛选出EMEA
                                            '删除
                                            
                                            With EmailReadyToSend.GetInspector.WordEditor
                                                .tables(t).rows(i).Delete
                                            End With
                                        End If
                                    Next
                                End If
                            Next
                            '----/EMEA-----
                            
                            
                            
                            

                            'MsgBox ("catched NA at : " & EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(i).Cells(j))
                            'MsgBox (t)  't 可以正确迭代
                            t = t + 1
                        Loop While t < EmailReadyToSend.GetInspector.WordEditor.tables.Count + 1
                        '-----------/ LA NA EMEA----------------
                        
                        '--------------删除不要的小表------------
                        'MsgBox ("t = " & t)
                        t = 1
                        Do
                            If InStr(EmailReadyToSend.GetInspector.WordEditor.tables(t).rows(1).Cells(1), "TOWER") Then
                                With EmailReadyToSend.GetInspector.WordEditor
                                    .tables(t).Delete
                                End With
                            End If
                            t = t + 1
                        Loop While t < EmailReadyToSend.GetInspector.WordEditor.tables.Count + 1
                        '----------/小表------------
                        

                        '===================/处理邮件中的表格==============================
                    
                        
                        EmailReadyToSend.Display  '变成草稿展示出来。
                        
                        'EmailReadyToSend.Send     '草稿直接转发出去

                    End If
                End If
            Next
        End If
    Next
    'clean up
    Set EmailReadyToSend = Nothing
    Set FwEmail = Nothing
    'objInbox Nothing
    'subFolder = Nothing
    'FwFromWho = Nothing
    'FwFromTo = Nothing
    'FwEmailDate = Nothing
    
End Sub

