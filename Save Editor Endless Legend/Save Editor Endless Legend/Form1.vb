Imports System.IO
Imports System.IO.Compression
Imports System
Imports System.Management
Imports System.Xml




Public Class Form1
    Dim save As String
    Dim xmlsave As New XmlDocument
    Dim nodo As XmlNode
    Dim xmlPath As String
    Dim result As DialogResult
    Dim extractPath As String
    Dim TextBox() As TextBox = {TextBox1, TextBox2, TextBox3, TextBox4, TextBox5, TextBox6,
                                TextBox7, TextBox8, TextBox9, TextBox10, TextBox11, TextBox12,
                                TextBox13, TextBox14, TextBox15, TextBox16, TextBox17, TextBox18,
                                TextBox19, TextBox20, TextBox21, TextBox22, TextBox23}

    Private Sub AbrirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AbrirToolStripMenuItem.Click
        result = OpenFileDialog1.ShowDialog()

        If result = DialogResult.OK Then
            save = OpenFileDialog1.FileName
            Label1.Text = String.Format("Editando: {0}", save)
            ProgressBar1.Value = 0
            extractPath = String.Format("{0}\Edit", Path.GetDirectoryName(save))
            If System.IO.Directory.Exists(extractPath) Then

                For Each s As String In My.Computer.FileSystem.GetFiles(extractPath + "\Atlas")
                    My.Computer.FileSystem.DeleteFile(s)
                Next
                For Each s As String In My.Computer.FileSystem.GetFiles(extractPath + "\Endless Legend")
                    My.Computer.FileSystem.DeleteFile(s)
                Next
                For Each s As String In My.Computer.FileSystem.GetFiles(extractPath)
                    My.Computer.FileSystem.DeleteFile(s)
                Next
                My.Computer.FileSystem.DeleteDirectory(extractPath + "\Atlas", FileIO.DeleteDirectoryOption.DeleteAllContents)
                My.Computer.FileSystem.DeleteDirectory(extractPath + "\Endless Legend", FileIO.DeleteDirectoryOption.DeleteAllContents)
                My.Computer.FileSystem.DeleteDirectory(extractPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

            End If
            ZipFile.ExtractToDirectory(save, extractPath)

            ProgressBar1.Value = 20

            xmlPath = String.Format("{0}\Endless Legend\Game.xml", extractPath)
            xmlsave.Load(xmlPath)
            ProgressBar1.Value = 40

            'Polvo e Influencia
            nodo = xmlsave.SelectSingleNode("//Property[@Name='BankAccount']")
            TextBox1.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='EmpirePointStock']")
            TextBox2.Text = String.Format(nodo.InnerText)
            ProgressBar1.Value = 60

            'Recursos Estratégicos
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic1Stock']")
            TextBox3.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic2Stock']")
            TextBox4.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic3Stock']")
            TextBox5.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic4Stock']")
            TextBox6.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic5Stock']")
            TextBox7.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic6Stock']")
            TextBox8.Text = String.Format(nodo.InnerText)
            ProgressBar1.Value = 80

            'Recursos Estratégicos
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury1Stock']")
            TextBox14.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury2Stock']")
            TextBox13.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury3Stock']")
            TextBox12.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury4Stock']")
            TextBox11.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury5Stock']")
            TextBox10.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury6Stock']")
            TextBox9.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury7Stock']")
            TextBox19.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury8Stock']")
            TextBox18.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury9Stock']")
            TextBox17.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury10Stock']")
            TextBox16.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury11Stock']")
            TextBox15.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury12Stock']")
            TextBox23.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury13Stock']")
            TextBox22.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury14Stock']")
            TextBox21.Text = String.Format(nodo.InnerText)
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury15Stock']")
            TextBox20.Text = String.Format(nodo.InnerText)
            ProgressBar1.Value = 100

            xmlsave.PreserveWhitespace = True
            ProgressBar1.Value = 0
        End If



    End Sub

    Private Sub GuardarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuardarToolStripMenuItem.Click

        result = SaveFileDialog1.ShowDialog()

        If result = DialogResult.OK Then
            ProgressBar1.Value = 0

            'Polvo e Influencia
            nodo = xmlsave.SelectSingleNode("//Property[@Name='BankAccount']")
            nodo.InnerText = TextBox1.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='EmpirePointStock']")
            nodo.InnerText = TextBox2.Text
            ProgressBar1.Value = 20

            'Recursos Estratégicos
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic1Stock']")
            nodo.InnerText = TextBox3.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic2Stock']")
            nodo.InnerText = TextBox4.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic3Stock']")
            nodo.InnerText = TextBox5.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic4Stock']")
            nodo.InnerText = TextBox6.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic5Stock']")
            nodo.InnerText = TextBox7.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Strategic6Stock']")
            nodo.InnerText = TextBox8.Text
            ProgressBar1.Value = 40

            'Recursos Estratégicos
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury1Stock']")
            nodo.InnerText = TextBox14.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury2Stock']")
            nodo.InnerText = TextBox13.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury3Stock']")
            nodo.InnerText = TextBox12.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury4Stock']")
            nodo.InnerText = TextBox11.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury5Stock']")
            nodo.InnerText = TextBox10.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury6Stock']")
            nodo.InnerText = TextBox9.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury7Stock']")
            nodo.InnerText = TextBox19.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury8Stock']")
            nodo.InnerText = TextBox18.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury9Stock']")
            nodo.InnerText = TextBox17.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury10Stock']")
            nodo.InnerText = TextBox16.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury11Stock']")
            nodo.InnerText = TextBox15.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury12Stock']")
            nodo.InnerText = TextBox23.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury13Stock']")
            nodo.InnerText = TextBox22.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury14Stock']")
            nodo.InnerText = TextBox21.Text
            nodo = xmlsave.SelectSingleNode("//Property[@Name='Luxury15Stock']")
            nodo.InnerText = TextBox20.Text
            ProgressBar1.Value = 60

            xmlsave.Save(xmlPath)
            save = SaveFileDialog1.FileName
            ProgressBar1.Value = 80

            extractPath = String.Format("{0}\Edit", Path.GetDirectoryName(save))

            xmlPath = String.Format("{0}\Endless Legend\GameSaveDescriptor.xml", extractPath)
            xmlsave.Load(xmlPath)
            nodo = xmlsave.SelectSingleNode("//Title")
            Dim Extension() As Char = {".", "z", "i", "p"}
            Dim nom As String = save
            nom = Path.GetFileName(nom)
            nom = nom.TrimEnd(Extension)
            nodo.InnerText = String.Format(nom)
            xmlsave.Save(xmlPath)

            ZipFile.CreateFromDirectory(extractPath, save)

            For Each s As String In My.Computer.FileSystem.GetFiles(extractPath + "\Atlas")
                My.Computer.FileSystem.DeleteFile(s)
            Next
            For Each s As String In My.Computer.FileSystem.GetFiles(extractPath + "\Endless Legend")
                My.Computer.FileSystem.DeleteFile(s)
            Next
            For Each s As String In My.Computer.FileSystem.GetFiles(extractPath)
                My.Computer.FileSystem.DeleteFile(s)
            Next
            My.Computer.FileSystem.DeleteDirectory(extractPath + "\Atlas", FileIO.DeleteDirectoryOption.DeleteAllContents)
            My.Computer.FileSystem.DeleteDirectory(extractPath + "\Endless Legend", FileIO.DeleteDirectoryOption.DeleteAllContents)
            My.Computer.FileSystem.DeleteDirectory(extractPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

            ProgressBar1.Value = 100
        End If
    End Sub

    Private Sub CréditosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CréditosToolStripMenuItem.Click
        Process.Start("http://steamcommunity.com/id/Dr_Bonsai")
    End Sub

    Private Sub AyudaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AyudaToolStripMenuItem.Click
        MessageBox.Show("Este Save Editor es muy simple de usar. En Archivo>Cargar Selecciona el archivo de guardado que quieras editar de tu carpeta Documentos\Endless Legend\Save Files y pulsa Abrir. Una vez abierto, cambia los valores y guárdalo con el nombre que quieras.", "Información", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox14_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox14.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox15_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox15.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox16_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox16.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox17_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox17.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox18_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox18.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox19_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox19.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox20_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox20.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox21_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox21.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox22_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox22.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
    Private Sub TextBox23_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox23.KeyPress
        If InStr(1, "0123456789.-" & Chr(8), e.KeyChar) = 0 Then
            e.KeyChar = ""
        End If
    End Sub
End Class
