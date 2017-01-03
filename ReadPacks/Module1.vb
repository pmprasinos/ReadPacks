
Module Module1
    Dim objConn As New System.Data.SqlClient.SqlConnection
    Dim cmd As New System.Data.SqlClient.SqlCommand
    Dim UnknownDir As String = "\\slfs01\shared\Doc-Pac-Ship\PackScan\UNKNOWN"
    Sub Main()
        objConn.ConnectionString = "Server=SLREPORT01; Database=WFLocal; persist security info=False; trusted_connection=Yes;"
        cmd.Connection = objConn



        filenamedpacks()
        Threading.Thread.Sleep(100)


        Dim UserPath As String = "\\slfs01\shared\prasinos\packs"
        Dim TargetFile As String = "-"

        UserPath = "\\slfs01\shared\prasinos\packs\"

        If UserPath <> "" Then
            Dim bc As New Bytescout.BarCodeReader.Reader()
            bc.BarcodeTypesToFind.Code39 = True
            bc.BarcodeTypesToFind.Code39Extended = True
            ' bc.MaxNumberOfBarcodesPerDocument = 100

            objConn.Open()
            Dim p As Integer = 0
            For Each f In FileIO.FileSystem.GetFiles("\\slfs01\shared\Doc-Pac-Ship\packscan", FileIO.SearchOption.SearchAllSubDirectories, "*.pdf")
                p = p + 1
                Dim invnum As String = Split(Dir(f), "__")(0)
                Debug.Print(invnum)
                cmd.CommandText = " merge wflocal..SHIPMENTS
                                    USING (SELECT @INVOICE_NO AS INVOICE_NO) AS INVOICE 
                                    ON INVOICE.INVOICE_NO=SHIPMENTS.INVOICE_NO 
                                        WHEN NOT MATCHED THEN
	                                    INSERT(INVOICE_NO, SCANNED, COMPANYNAME) VALUES(@INVOICE_NO, 1, @PATH)
                                    WHEN MATCHED THEN
	                                    UPDATE  SET SCANNED = 0;"
                cmd.Parameters.Clear()
                cmd.Parameters.AddWithValue("@INVOICE_NO", invnum)
                cmd.Parameters.AddWithValue("@PATH", f)
                cmd.ExecuteNonQuery()
            Next
            '   Try
            If Dir(UserPath & "*") <> "" Then
                'Threading.Thread.Sleep(150)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Console.Write(Dir(UserPath) & "...")
                Dim FileCount As Integer = FileIO.FileSystem.GetFiles(UserPath, FileIO.SearchOption.SearchTopLevelOnly, {"*.pdf"}).Count
                Console.WriteLine(FileCount & " UNKNOWN PACKS IN " & UserPath)
                Dim u As Integer = 0


                Dim foundinvoices As String = ""
                For Each file In FileIO.FileSystem.GetFiles(UserPath, FileIO.SearchOption.SearchTopLevelOnly)
                    u = u + 1
                    Console.Write("File " & u & " of " & FileCount & ">>" & Replace(Dir(file), ".pdf", "") & "...")
                    If InStr(file, ".pdf", CompareMethod.Text) > 0 Then
                        TargetFile = file

                        Dim ChopList As New List(Of Integer)
                        Dim InvList As New List(Of String)

                        Dim Barcodes As Bytescout.BarCodeReader.FoundBarcode() = bc.ReadFrom(TargetFile)
                        Dim LastPage As Integer = Barcodes.Last.Page + 1
                        Dim barcst As String = ""
                        For bcnum = 0 To Barcodes.Count - 1

                            Dim d As Integer = Array.IndexOf(Barcodes, Barcodes(bcnum))
                            '  ChopList.Add(barc)

                            If Not ChopList.Contains((Barcodes(bcnum).Page)) Then
                                ChopList.Add(Barcodes(bcnum).Page)
                                InvList.Add(Split(Barcodes(bcnum).Value, "(")(0))
                            End If
                        Next

                        For c = 0 To ChopList.Count - 1

                            Dim bcnum As Integer = ChopList(c)
                            Dim barc As String
                            barc = InvList(c)
                            Debug.Print(barc)



                            barcst = barcst & barc
                            cmd.CommandText = "select * from wflocal.dbo.shipments where invoice_no = @INVOICE_NO"
                            cmd.Parameters.Clear()
                            cmd.Parameters.AddWithValue("@INVOICE_NO", barc)
                            Using dr As SqlClient.SqlDataReader = cmd.ExecuteReader
                                Dim OutPDF As String
                                If dr.HasRows Then
                                    dr.Read()
                                    OutPDF = "S:\Doc-Pac-Ship\PackScan\" & dr("PARTNO").ToString & "\"
                                    If Not FileIO.FileSystem.DirectoryExists(OutPDF) Then FileIO.FileSystem.CreateDirectory(OutPDF)
                                    Dim f As String = Replace(Split(dr("SHIPPED_DTIME").ToString, " ")(0), "/", "-")
                                    OutPDF = OutPDF & barc & "__" & f & ".pdf"
                                Else
                                    OutPDF = "\\slfs01\shared\Doc-Pac-Ship\PackScan\UNKNOWN\PACK.pdf"
                                    foundinvoices = foundinvoices & barc & " NOTFOUND  " & vbCrLf
                                End If

                                Dim PN As Integer = 0
                                Do While FileIO.FileSystem.FileExists(OutPDF)
                                    PN = PN + 1
                                    If Not FileIO.FileSystem.FileExists(Replace(OutPDF, ".pdf", "(" & PN & ").pdf")) Then
                                        OutPDF = Replace(OutPDF, ".pdf", "(" & PN & ").pdf")
                                    End If
                                Loop

                                Dim pagelist As New List(Of Integer)

                                PN = 0

                                If c = ChopList.Count - 1 Then
                                    pagelist.Add(ChopList(c) - 1)
                                Else

                                    For page As Integer = ChopList(c) To ChopList(c + 1) - 1
                                        pagelist.Add(page + 1)
                                    Next

                                End If
                                'pagelist.Add(Barcodes(bcnum).Page + 1) : pagelist.Add(Barcodes(bcnum).Page)
                                PdfManipulation2.ExtractPdfPages(file, pagelist, OutPDF)
                                Dim Pages As String = ""
                                For Each PAGE In pagelist
                                    Pages = Pages & PAGE & "."
                                Next
                                Console.WriteLine("PAGES (" & Pages & ") MOVED TO " & OutPDF)
                                foundinvoices = foundinvoices & barc & "    FOUND  " & InvList(c) & vbCrLf
                            End Using

                            cmd.CommandText = "UPDATE WFLOCAL..SHIPMENTS SET SCANNED = 1 WHERE INVOICE_NO = @INVOICE_NO"
                            cmd.ExecuteNonQuery()


                        Next

                    End If
                Next
                FileIO.FileSystem.WriteAllText(Replace(TargetFile, ".pdf", ".txt", 1, -1, CompareMethod.Text), foundinvoices, False)

            End If
            '  Catch
            '  Finally
            objConn.Close()
            ' End Try
        End If

    End Sub

    Private Sub filenamedpacks()
        Dim FileCount As Integer = FileIO.FileSystem.GetFiles("\\slfs01\shared\Doc-Pac-Ship\PackScan\UNKNOWN").Count
        Console.WriteLine(FileCount & " UNKNOWN PACKS IN \\slfs01\shared\Doc-Pac-Ship\PackScan\UNKNOWN\")
        Dim u As Integer = 0
        For Each file In FileIO.FileSystem.GetFiles("\\slfs01\shared\Doc-Pac-Ship\PackScan\UNKNOWN")
            u = u + 1
            Dim FILENAME As String = Replace(Dir(file), ".pdf", "")

            Console.Write("File " & u & " of " & FileCount & ">>" & FILENAME & "...")

            If Not InStr(file, "PACK") > 0 Then
                objConn.Open()
                cmd.CommandText = "select * from wflocal.dbo.shipments where invoice_no = @INVOICE_NO"
                cmd.Parameters.Clear()
                cmd.Parameters.AddWithValue("@INVOICE_NO", Replace(FILENAME, ".pdf", ""))
                Dim OutPDF As String = ""
                Dim f As String = ""
                Using dr As SqlClient.SqlDataReader = cmd.ExecuteReader
                    If dr.HasRows Then
                        dr.Read()
                        OutPDF = "S:\Doc-Pac-Ship\PackScan\" & dr("PARTNO").ToString & "\"
                        If Not FileIO.FileSystem.DirectoryExists(OutPDF) Then FileIO.FileSystem.CreateDirectory(OutPDF)
                        f = Replace(Split(dr("SHIPPED_DTIME").ToString, " ")(0), "/", "-")
                    Else
                        Console.WriteLine("NOT FILED, NO MATCH FOUND")
                    End If
                End Using

                If OutPDF = "" Then
                    cmd.CommandText = "UPDATE WFLOCAL..SHIPMENTS SET SCANNED = 1 WHERE INVOICE_NO = @INVOICE_NO"
                    cmd.ExecuteNonQuery()
                End If
                objConn.Close()

                OutPDF = OutPDF & Replace(FILENAME, ".pdf", "") & "__" & f & ".pdf"
                        Dim PN As Integer = 0
                        Do While FileIO.FileSystem.FileExists(OutPDF)
                            PN = PN + 1
                            If Not FileIO.FileSystem.FileExists(Replace(OutPDF, ".pdf", "(" & PN & ").pdf")) Then
                                OutPDF = Replace(OutPDF, ".pdf", "(" & PN & ").pdf")
                            End If
                        Loop
                FileIO.FileSystem.MoveFile(file, OutPDF)
                Console.WindowWidth = 150
                Console.WriteLine("MOVED TO " & OutPDF)
                        PN = 0







            Else
                Console.WriteLine("PLEASE REVIEW. FILE HAS NOT BEEN RENAMED")
            End If

        Next file
    End Sub

End Module
