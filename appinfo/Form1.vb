'This NotePad is created by Mr.A studio
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ToolStripStatusLabel1.Text = DateTime.Today
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub NewToolStripButton_Click(sender As Object, e As EventArgs) Handles NewToolStripButton.Click
        If RichTextBox1.Modified Then
            Dim a As MsgBoxResult
            a = MsgBox("Do you want to save changes...", MsgBoxStyle.YesNoCancel, "New document ")
            If a = MsgBoxResult.No Then
                RichTextBox1.Clear()
            ElseIf a = MsgBoxResult.Cancel Then
            ElseIf a = MsgBoxResult.Yes Then
                SaveFileDialog1.ShowDialog()
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
                RichTextBox1.Clear()
            End If
        Else
            RichTextBox1.Clear()
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub

    Private Sub Font_StyleToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles Font_StyleToolStripMenuItem2.Click
        FontDialog1.ShowDialog()
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub Font_ColorToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles Font_ColorToolStripMenuItem2.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub BackgroundToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles BackgroundToolStripMenuItem2.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.BackColor = ColorDialog1.Color
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        SaveFileDialog1.ShowDialog()
        If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) Then
            Dim a As MsgBoxResult
            a = MsgBox("File is already exist,Would you like to replace it", MsgBoxStyle.YesNoCancel, "File exist")
            If a = MsgBoxResult.No Then
                SaveFileDialog1.Filter = "text file (.txt) | .txt"
                SaveFileDialog1.ShowDialog()
            ElseIf a = MsgBoxResult.Yes Then
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, True)
            End If
        Else
            SaveFileDialog1.Filter = "text file (.txt) | .txt"
            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, True)

            End If
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveFileDialog1.ShowDialog()
        If My.Computer.FileSystem.FileExists(SaveFileDialog1.FileName) Then
            Dim a As MsgBoxResult
            a = MsgBox("File is already exist,Would you like to replace it", MsgBoxStyle.YesNoCancel, "File exist")
            If a = MsgBoxResult.No Then
                SaveFileDialog1.ShowDialog()
            ElseIf a = MsgBoxResult.Yes Then
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
            End If
        Else
            Try
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If RichTextBox1.Modified Then
            Dim a As MsgBoxResult
            a = MsgBox("Do you want to save changes...", MsgBoxStyle.YesNoCancel, "Open document ")
            If a = MsgBoxResult.No Then
                OpenFileDialog1.ShowDialog()
                RichTextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            ElseIf a = MsgBoxResult.Cancel Then
            ElseIf a = MsgBoxResult.Yes Then
                My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
                RichTextBox1.Clear()
            End If
        Else
            OpenFileDialog1.ShowDialog()
            Try
                RichTextBox1.Text = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.ShowDialog()
    End Sub
    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        Dim a, b As String
        a = InputBox("Enter text to be search")
        b = InStr(RichTextBox1.Text, a)
        If b Then
            RichTextBox1.Focus()
            RichTextBox1.SelectionStart = b - 1
            RichTextBox1.SelectionLength = Len(a)
        Else
            MsgBox("Text is not found :(")
        End If
    End Sub

    Private Sub Zoom_InToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles Zoom_InToolStripMenuItem2.Click
        RichTextBox1.Font = New Font(RichTextBox1.Font.FontFamily, RichTextBox1.Font.Size + 3, RichTextBox1.Font.Style)
    End Sub

    Private Sub Zoom_OutToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles Zoom_OutToolStripMenuItem3.Click
        RichTextBox1.Font = New Font(RichTextBox1.Font.FontFamily, RichTextBox1.Font.Size - 3, RichTextBox1.Font.Style)
    End Sub

    Private Sub Restore_Default_ZoomToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles Restore_Default_ZoomToolStripMenuItem4.Click
        RichTextBox1.Font = New Font("", 9)
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("This NotePad is created by Mr.A studio", MsgBoxStyle.OkOnly)
    End Sub
End Class
