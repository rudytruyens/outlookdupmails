Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Security.Cryptography

Public Class Form1
    Dim outlookapp As Outlook.Application = New Outlook.Application
    Dim outlooknamespace As Outlook.NameSpace = outlookapp.GetNamespace("MAPI")
    Dim cn As OleDbConnection = New OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source=" + Application.StartupPath + "\dupmails.mdb")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim answer As MsgBoxResult = MsgBox(lblsummary.Text & vbCrLf & "Continue ?", MsgBoxStyle.OkCancel)
        If answer = MsgBoxResult.Ok Then
            cn.Open()
            Dim cmd = New OleDbCommand()
            cmd.CommandText = "DELETE mails.* FROM mails"
            cmd.Connection = cn
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            'Dim outlookinboxfolder As Outlook.MAPIFolder = outlooknamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
            'Dim outlookinboxfldr As Outlook.Items = outlookinboxfolder.Items
            Dim root As Outlook.MAPIFolder = outlooknamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
            Dim fldrs As Outlook.Folders = root.Folders
            Dim folderitems As Outlook.Items

            For Each fol As Outlook.MAPIFolder In fldrs
                If fol.Name.ToLower = "inbox" Then
                    'Process inbox
                    processfolder(fol.Items, fol.Name, fol.Items.Count)
                    'process inbox subfolders
                    If ListBox1.Visible = True Then
                        Dim inbox As Outlook.Folders = fol.Folders
                        For Each inboxfolder As Outlook.MAPIFolder In inbox
                            If ListBox1.SelectedItems.Contains(inboxfolder.Name) Then
                                folderitems = inboxfolder.Items
                                Dim total As String = folderitems.Count.ToString


                                processfolder(folderitems, inboxfolder.Name, folderitems.Count.ToString)

                            End If
                        Next
                    End If
                End If
            Next



            lblstatus.Text = "Analyzing..."
            Application.DoEvents()
            cmd = New OleDbCommand()
            cmd.Connection = cn
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE tokeep.* FROM tokeep"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "INSERT INTO Tokeep ( FirstOfEntryid, Hash, CountOfHash ) Select First(Mails.Entryid) As FirstOfEntryid, Mails.Hash, Count(Mails.Hash) As CountOfHash From Mails Group By Mails.Hash HAVING(((Count(Mails.Hash)) > 1))"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE doubles.* FROM doubles"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "INSERT INTO Doubles ( Entryid, Hash, [Size] ) Select Mails.Entryid, Mails.Hash, Mails.Size From tokeep INNER Join Mails On tokeep.Hash = Mails.Hash Order By Mails.Hash"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE Doubles.*, Doubles.Entryid From Doubles Where (((Doubles.Entryid) In (Select firstofentryid From tokeep)))"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "Select * From Doubles"
            Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim ds As DataSet = New DataSet()
            da.Fill(ds)
            da.Fill(ds, "Table1")
            lblstatus.Text = "Moving to trash folder..."
            Application.DoEvents()
            'Dim trash As Outlook.MAPIFolder
            'trash = outlooknamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)
            Dim outlookmessage As Outlook.MailItem
            If ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    outlookmessage = outlooknamespace.GetItemFromID(row.Item("Entryid"))
                    outlookmessage.Delete()
                Next
            Else
                MsgBox("no double email records found")
            End If
            lblstatus.Text = "Done !"
            cmd.Dispose()
            cn.Close()

        End If
    End Sub
    Private Sub processfolder(ByRef folderitems As Outlook.Items, foldername As String, total As String)
        Dim outlookmessage As Outlook.MailItem
        Dim cmd As OleDbCommand
        Dim plaintext As String
        For f As Integer = 1 To folderitems.Count
            If folderitems.Item(f).MessageClass = "IPM.Note" Then
                outlookmessage = folderitems.Item(f)
                If outlookmessage.Body <> Nothing Or outlookmessage.Attachments.Count <> 0 Then
                    cmd = New OleDbCommand()
                    cmd.Connection = cn
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandText = "insertmails"
                    cmd.Parameters.Add("@Entryid", OleDbType.VarChar).Value = outlookmessage.EntryID
                    plaintext = outlookmessage.Body
                    If chkSameSentDate.Checked = True Then
                        plaintext += outlookmessage.SentOn.Ticks.ToString
                    End If
                    If chkSameSize.Checked = True Then
                        plaintext += outlookmessage.Size.ToString
                    End If
                    cmd.Parameters.Add("@hash", OleDbType.VarChar).Value = GenerateSHA256String(plaintext)  ' Add Parameter
                    cmd.Parameters.Add("@thissize", OleDbType.VarChar).Value = outlookmessage.Size
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                End If

            End If

            lblstatus.Text = "Reading " + foldername + " " + f.ToString + "/" + total
        Next
    End Sub
    Private Function GenerateSHA256String(inputString As String) As String
        Dim sha256 = SHA256Managed.Create()
        Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(inputString)
        Dim hash As Byte() = sha256.ComputeHash(bytes)
        Return Convert.ToBase64String(hash)
    End Function
    Private Sub Setsummary()
        Dim Summary As String
        If chkSameSentDate.Checked = True Then
            Summary = "Emails with the same contents and sent on the same date will be moved to the 'deleted items' folder"
            If chkSameSize.Checked = True Then
                Summary = "Emails with the same contents and sent on the same date and having exactly the same message size will be moved to the 'deleted items' folder"
            End If
        Else
            If chkSameSize.Checked = True Then
                Summary = "Emails with the same contents and having exactly the same message size will be moved to the 'deleted items' folder"
            Else
                Summary = "Emails with the same contents will be moved to the 'deleted items' folder (even if the size and the date send of these messages is different)."
            End If
        End If
        lblsummary.Text = Summary
    End Sub


    Private Sub chkSameSentDate_CheckedChanged(sender As Object, e As EventArgs) Handles chkSameSentDate.CheckedChanged
        Setsummary()
    End Sub

    Private Sub chkSameSize_CheckedChanged(sender As Object, e As EventArgs) Handles chkSameSize.CheckedChanged
        Setsummary()
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim root As Outlook.MAPIFolder = outlooknamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
        Dim fldrs As Outlook.Folders = root.Folders
        For Each fol As Outlook.MAPIFolder In fldrs
            If fol.Name.ToLower = "inbox" Then
                Dim inbox As Outlook.Folders = fol.Folders
                For Each inboxfolder As Outlook.MAPIFolder In inbox
                    If inboxfolder.Name.ToLower <> "deleted items" And inboxfolder.Name.ToLower <> "trash" Then
                        ListBox1.Items.Add(inboxfolder.Name)
                    End If
                Next
            End If
        Next
        If ListBox1.Items.Count > 0 Then
            ListBox1.Visible = True
            lblsubfolders.Visible = True
            Me.Height = 485
        Else
            ListBox1.Visible = False
            lblsubfolders.Visible = False
            Me.Height = 300
        End If
        lblsubfolders.Text = "The inbox folder contains subfolders. Which of the subfolders do you want to process ? " & vbCrLf & "Multiple items can be selected:"
    End Sub


    Private Sub Form1_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        outlookapp = Nothing
        outlooknamespace = Nothing
    End Sub
End Class
