Imports System.IO
Imports System.Xml

Public Class Form1

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

        ' 遍尋路徑下所有目錄及檔案
        Dim strSource As New System.IO.DirectoryInfo(txtSource.Text) '設定路徑變數並實作
        Dim strMbsaXml As System.IO.FileInfo    ' 宣告檔案名稱變數
        Dim subDir As System.IO.DirectoryInfo   ' 宣告子目錄變數

        '清除路徑頭尾的空白
        txtTarget.Text = Trim(txtTarget.Text)

        '建立結果XML檔
        '如果目的路徑為空值，則使用來源路徑
        If txtTarget.Text = "" Then
            txtTarget.Text = txtSource.Text
        End If

        '判斷來源目的後面是否有"\"
        If txtTarget.Text.EndsWith("\") = False Then
            txtTarget.Text += "\"
        End If

        Dim strDataTime As String = DateTime.Now.ToString("yyyyMMddHHmmss") '取目前的日期時間作為檔名的一部份，避免檔名重複

        Dim strXmlFile As String = txtTarget.Text + "MBSA_" + strDataTime + ".x_l"  ' 設定檔案名稱(含路徑)
        Dim strCsvFile As String = txtTarget.Text + "MBSA_" + strDataTime + ".csv"

        Dim fileMbsaXml As System.IO.StreamWriter

        fileMbsaXml = My.Computer.FileSystem.OpenTextFileWriter(strXmlFile, True)
        fileMbsaXml.WriteLine("<XMLOut>")  '寫入ROOT節點
        fileMbsaXml.Close()

        Dim fileMbsaCsv As System.IO.StreamWriter   ' 新增CSV檔案部分
        fileMbsaCsv = My.Computer.FileSystem.OpenTextFileWriter(strCsvFile, True)   ' 新增CSV檔案部分
        fileMbsaCsv.WriteLine("NampIP,UpdateData")  ' 新增CSV檔案部分

        ' 設定暫存檢查結果變數
        Dim tmpTXT

        '搜尋子目錄下所有的XML檔案
        For Each strMbsaXml In strSource.GetFiles("*.xml", IO.SearchOption.AllDirectories)
            '取出檔案的名稱(去除"MBSA_"，以作為主機名稱辨識
            txtKB_Installed.Text = txtKB_Installed.Text + Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + vbCrLf
            tmpTXT = "" ' 每次新的檔案就將暫存變數清空

            Dim rdMBSAXML As XmlReader = XmlReader.Create(strMbsaXml.FullName) '讀取MBSA XML檔案

            Try
                ' 為針對不完整XML檔案，抓取錯誤，並將資訊填入輸出檔，以判明和台主機MBSA檔案毀損
                Do While rdMBSAXML.Read()  '讀取XML中所有節點
                    '將MBSA XML中所有 UpData 元素讀取出來，以判斷是否安裝及等級
                    '元素 UpData 中，是否安裝屬性為第5個(由0開始計算)
                    ' KB 編號屬性為第3個，等級為第6個
                    If rdMBSAXML.NodeType = XmlNodeType.Element AndAlso rdMBSAXML.Name = "UpdateData" Then
                        '判斷更新是否安裝，且等級大於等於3
                        If rdMBSAXML.GetAttribute(5) = "false" And rdMBSAXML.GetAttribute(6) >= 3 Then
                            ' tmpTXT = rdMBSAXML.GetAttribute(3) + " / " + rdMBSAXML.GetAttribute(5) + " / " + rdMBSAXML.GetAttribute(6) + vbCrLf
                            Dim intComp As Integer = tmpTXT.IndexOf(rdMBSAXML.GetAttribute(3))
                            If intComp = -1 Then
                                tmpTXT += "KB" + rdMBSAXML.GetAttribute(3) + vbCrLf
                            End If
                        End If
                    End If
                Loop

            Catch ex As Exception
                fileMbsaXml = My.Computer.FileSystem.OpenTextFileWriter(strXmlFile, True)
                fileMbsaXml.WriteLine("<MBSA_Check>")  '包住每個主機的結果
                fileMbsaXml.WriteLine("  <NameIp>" + Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + "</NameIp>") '寫入主機名稱(含IP)
                fileMbsaXml.WriteLine("  <UpdataData>此主機MBSA檔案毀損，無法正確開啟</UpdataData>")
                fileMbsaXml.WriteLine("</MBSA_Check>")
                fileMbsaXml.Close()
                fileMbsaCsv.WriteLine(Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + ",此主機MBSA檔案毀損，無法正確開啟") ' 新增CSV檔案部分
                txtKB_Installed.Text = txtKB_Installed.Text + "此主機MBSA檔案毀損，無法正確開啟" + vbCrLf + "------------" + vbCrLf
                Continue For
            End Try

            rdMBSAXML.Close() '關閉XML檔案
            '將有結果的部分放入預覽框中
            txtKB_Installed.Text = txtKB_Installed.Text + tmpTXT + "------------" + vbCrLf

            '將結果寫入檔案
            fileMbsaXml = My.Computer.FileSystem.OpenTextFileWriter(strXmlFile, True)
            fileMbsaXml.WriteLine("<MBSA_Check>")  '包住每個主機的結果
            'fileMbsaXml.WriteLine("  <NameIp>" + subDir.Name.Replace("_REPORT", "") + "</NameIp>") '寫入主機名稱(含IP)
            fileMbsaXml.WriteLine("  <NameIp>" + Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + "</NameIp>") '寫入主機名稱(含IP)

            Dim tmpCsvTXT ' 因應新增CSV檔案部分，增加此暫存變數存放CSV內容

            ' 判斷tmpTXT有無資料，決定寫入檔案內容
            If tmpTXT = "" Then
                fileMbsaXml.WriteLine("  <UpdataData>微軟(KB) : 無未更新</UpdataData>")
                tmpCsvTXT = "微軟(KB) : 無未更新"   ' 新增CSV檔案部分
            Else
                fileMbsaXml.WriteLine("    <UpdataData> 微軟(KB):" + vbCrLf + tmpTXT + "未完成更新</UpdataData>")
                tmpCsvTXT = """微軟(KB):" + vbCrLf + tmpTXT + "未完成更新"""    ' 新增CSV檔案部分
            End If
            fileMbsaXml.WriteLine("</MBSA_Check>")
            fileMbsaXml.Close()

            fileMbsaCsv.WriteLine(Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + "," + tmpCsvTXT) ' 新增CSV檔案部分
        Next



        '尋找所有子目錄
        ' For Each subDir In strSource.GetDirectories

        '取出子目錄的名稱(去除"_REPORT"，以作為主機名稱辨識
        'txtKB_Installed.Text = txtKB_Installed.Text + subDir.Name.Replace("_REPORT", "") + vbCrLf

        ''搜尋子目錄下所有的XML檔案
        'For Each strMbsaXml In subDir.GetFiles("*.xml")
        '    '取出檔案的名稱(去除"MBSA_"，以作為主機名稱辨識
        '    txtKB_Installed.Text = txtKB_Installed.Text + Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + vbCrLf
        '    tmpTXT = "" ' 每次新的檔案就將暫存變數清空

        '    Dim rdMBSAXML As XmlReader = XmlReader.Create(strMbsaXml.FullName) '讀取MBSA XML檔案

        '    Try
        '        ' 為針對不完整XML檔案，抓取錯誤，並將資訊填入輸出檔，以判明和台主機MBSA檔案毀損
        '        Do While rdMBSAXML.Read()  '讀取XML中所有節點
        '            '將MBSA XML中所有 UpData 元素讀取出來，以判斷是否安裝及等級
        '            '元素 UpData 中，是否安裝屬性為第5個(由0開始計算)
        '            ' KB 編號屬性為第3個，等級為第6個
        '            If rdMBSAXML.NodeType = XmlNodeType.Element AndAlso rdMBSAXML.Name = "UpdateData" Then
        '                '判斷更新是否安裝，且等級大於等於3
        '                If rdMBSAXML.GetAttribute(5) = "false" And rdMBSAXML.GetAttribute(6) >= 3 Then
        '                    ' tmpTXT = rdMBSAXML.GetAttribute(3) + " / " + rdMBSAXML.GetAttribute(5) + " / " + rdMBSAXML.GetAttribute(6) + vbCrLf
        '                    tmpTXT += "KB" + rdMBSAXML.GetAttribute(3) + vbCrLf
        '                End If
        '            End If
        '        Loop

        '    Catch ex As Exception
        '        fileMbsaXml = My.Computer.FileSystem.OpenTextFileWriter(strXmlFile, True)
        '        fileMbsaXml.WriteLine("<MBSA_Check>")  '包住每個主機的結果
        '        fileMbsaXml.WriteLine("  <NameIp>" + Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + "</NameIp>") '寫入主機名稱(含IP)
        '        fileMbsaXml.WriteLine("  <UpdataData>此主機MBSA檔案毀損，無法正確開啟</UpdataData>")
        '        fileMbsaXml.WriteLine("</MBSA_Check>")
        '        fileMbsaXml.Close()
        '        fileMbsaCsv.WriteLine(Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + ",此主機MBSA檔案毀損，無法正確開啟") ' 新增CSV檔案部分
        '        txtKB_Installed.Text = txtKB_Installed.Text + "此主機MBSA檔案毀損，無法正確開啟" + vbCrLf + "------------" + vbCrLf
        '        Continue For
        '    End Try

        '    rdMBSAXML.Close() '關閉XML檔案
        '    '將有結果的部分放入預覽框中
        '    txtKB_Installed.Text = txtKB_Installed.Text + tmpTXT + "------------" + vbCrLf

        '    '將結果寫入檔案
        '    fileMbsaXml = My.Computer.FileSystem.OpenTextFileWriter(strXmlFile, True)
        '    fileMbsaXml.WriteLine("<MBSA_Check>")  '包住每個主機的結果
        '    'fileMbsaXml.WriteLine("  <NameIp>" + subDir.Name.Replace("_REPORT", "") + "</NameIp>") '寫入主機名稱(含IP)
        '    fileMbsaXml.WriteLine("  <NameIp>" + Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + "</NameIp>") '寫入主機名稱(含IP)

        '    Dim tmpCsvTXT ' 因應新增CSV檔案部分，增加此暫存變數存放CSV內容

        '    ' 判斷tmpTXT有無資料，決定寫入檔案內容
        '    If tmpTXT = "" Then
        '        fileMbsaXml.WriteLine("  <UpdataData>OS(含OFFICE更新) : 無嚴重度3級以上未更新</UpdataData>")
        '        tmpCsvTXT = "OS(含OFFICE更新) : 無嚴重度3級以上未更新"   ' 新增CSV檔案部分
        '    Else
        '        fileMbsaXml.WriteLine("    <UpdataData> OS(含OFFICE更新):" + vbCrLf + tmpTXT + "未完成更新</UpdataData>")
        '        tmpCsvTXT = """OS(含OFFICE更新):" + vbCrLf + tmpTXT + "未完成更新"""    ' 新增CSV檔案部分
        '    End If
        '    fileMbsaXml.WriteLine("</MBSA_Check>")
        '    fileMbsaXml.Close()

        '    fileMbsaCsv.WriteLine(Path.GetFileNameWithoutExtension(strMbsaXml.Name).Replace("MBSA_", "") + "," + tmpCsvTXT) ' 新增CSV檔案部分
        'Next

        'Next

        fileMbsaXml = My.Computer.FileSystem.OpenTextFileWriter(strXmlFile, True)
        fileMbsaXml.WriteLine("</XMLOut>")
        fileMbsaXml.Close()
        My.Computer.FileSystem.RenameFile(strXmlFile, "MBSA_" + strDataTime + ".xml")

        fileMbsaCsv.Close() ' 新增CSV檔案部分




        'Dim strMBSAXML As XmlReader = XmlReader.Create(txtSource.Text) '讀取MBSA XML檔案
        'Do While strMBSAXML.Read()  '讀取XML中所有節點
        '    Dim tmpTXT
        '    '將MBSA XML中所有 UpData 元素讀取出來，以判斷是否安裝及等級
        '    '元素 UpData 中，是否安裝屬性為第5個(由0開始計算)
        '    ' KB 編號屬性為第3個，等級為第6個
        '    If strMBSAXML.NodeType = XmlNodeType.Element AndAlso strMBSAXML.Name = "UpdateData" Then
        '        '判斷更新是否安裝，且等級大於等於3
        '        If strMBSAXML.GetAttribute(5) = "false" And strMBSAXML.GetAttribute(6) >= 3 Then
        '            tmpTXT = strMBSAXML.GetAttribute(3) + " / " + strMBSAXML.GetAttribute(5) + " / " + strMBSAXML.GetAttribute(6) + vbCrLf
        '            txtKB_Installed.Text = txtKB_Installed.Text + tmpTXT
        '        End If

        '    End If
        'Loop

        MessageBox.Show("OK")

    End Sub
End Class
