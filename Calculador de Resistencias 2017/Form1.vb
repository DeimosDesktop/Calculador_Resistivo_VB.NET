Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0
                Button1.BackColor = Color.Black
            Case 1
                TextBox1.Text = "1"
                Button1.BackColor = Color.Maroon
            Case 2
                TextBox1.Text = "2"
                Button1.BackColor = Color.Red
            Case 3
                TextBox1.Text = "3"
                Button1.BackColor = Color.Orange
            Case 4
                TextBox1.Text = "4"
                Button1.BackColor = Color.Yellow
            Case 5
                TextBox1.Text = "5"
                Button1.BackColor = Color.Green
            Case 6
                TextBox1.Text = "6"
                Button1.BackColor = Color.Blue
            Case 7
                TextBox1.Text = "7"
                Button1.BackColor = Color.Violet
            Case 8
                TextBox1.Text = "8"
                Button1.BackColor = Color.Gray
            Case 9
                TextBox1.Text = "9"
                Button1.BackColor = Color.White
        End Select
        ComboBox1.Enabled = False
        ComboBox2.Enabled = True
        ComboBox2.Focus()
    End Sub

    
    
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Select Case ComboBox2.SelectedIndex
            Case 0
                TextBox1.Text += "0"
                Button2.BackColor = Color.Black
            Case 1
                TextBox1.Text += "1"
                Button2.BackColor = Color.Maroon
            Case 2
                TextBox1.Text += "2"
                Button2.BackColor = Color.Red
            Case 3
                TextBox1.Text += "3"
                Button2.BackColor = Color.Orange
            Case 4
                TextBox1.Text += "4"
                Button2.BackColor = Color.Yellow
            Case 5
                TextBox1.Text += "5"
                Button2.BackColor = Color.Green
            Case 6
                TextBox1.Text += "6"
                Button2.BackColor = Color.Blue
            Case 7
                TextBox1.Text += "7"
                Button2.BackColor = Color.Violet
            Case 8
                TextBox1.Text += "8"
                Button2.BackColor = Color.Gray
            Case 9
                TextBox1.Text += "9"
                Button2.BackColor = Color.White

        End Select
        ComboBox2.Enabled = False
        ComboBox3.Enabled = True
        ComboBox3.Focus()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ComboBox1.Enabled = True
        ComboBox1.ResetText()
        ComboBox2.Enabled = True
        ComboBox2.ResetText()
        ComboBox3.Enabled = True
        ComboBox3.ResetText()
        ComboBox4.Enabled = True
        ComboBox4.ResetText()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        Button1.BackColor = Nothing
        Button2.BackColor = Nothing
        Button3.BackColor = Nothing
        Button4.BackColor = Nothing
        ListBox1.Items.Clear()
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False


    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        CheckBox1.Enabled = True
        Select Case ComboBox3.SelectedIndex
            Case 0
                Button3.BackColor = Color.Black
            Case 1
                TextBox1.Text += "0"
                Button3.BackColor = Color.Maroon
            Case 2
                TextBox1.Text += "00"
                Button3.BackColor = Color.Red
            Case 3
                TextBox1.Text += "000"
                Button3.BackColor = Color.Orange
            Case 4
                TextBox1.Text += "0000"
                Button3.BackColor = Color.Yellow
            Case 5
                TextBox1.Text += "00000"
                Button3.BackColor = Color.Green
            Case 6
                TextBox1.Text += "000000"
                Button3.BackColor = Color.Blue
            Case 7
                TextBox1.Text += "0000000"
                Button3.BackColor = Color.Violet
            Case 8
                TextBox1.Text = TextBox1.Text * 0.1
                Button3.BackColor = Color.Gray
            Case 9
                TextBox1.Text = TextBox1.Text * 0.01
                Button3.BackColor = Color.White

        End Select
        ComboBox3.Enabled = False
        ComboBox4.Enabled = True
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Select Case ComboBox4.SelectedIndex
            Case 0
                Button4.BackColor = Color.Red
                TextBox1.Text = " " & TextBox1.Text & " +- " & TextBox1.Text * 2 / 100 & "Ω"

            Case 1
                Button4.BackColor = Color.Gold
                TextBox1.Text = " " & TextBox1.Text & " +- " & TextBox1.Text * 5 / 100 & "Ω"
            Case 2
                Button4.BackColor = Color.Silver
                TextBox1.Text = " " & TextBox1.Text & " +- " & TextBox1.Text * 10 / 100 & "Ω"
        End Select
        ComboBox4.Enabled = False
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        GroupBox1.Enabled = True
        CheckBox1.Enabled = True
        CheckBox2.Enabled = True
        CheckBox3.Enabled = True
        Dim a = TextBox1.Text.IndexOf("")
        Dim b = TextBox1.Text.LastIndexOf("+-")
        Dim c = TextBox1.Text.Substring(a, b) - 1 + 1
        Dim medida3, resultado3, muestra3, medidafinal2 As String
        medida3 = TextBox1.Text.IndexOf("-") + 1
        medidafinal2 = TextBox1.Text.LastIndexOf("Ω")
        resultado3 = medidafinal2.ToString - medida3.ToString
        muestra3 = TextBox1.Text.Substring(medida3, resultado3)
        TextBox2.Text = muestra3.Remove(0, 1)
        TextBox3.Text = TextBox1.Text.Substring(a, b) - 1 + 1
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("https://deimoshackelectronica.blogspot.com.es/2016/08/codigo-de-colores-de-resistencias_18.html")
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        System.Diagnostics.Process.Start("https://mirpas.com")
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        Dim a As Integer = 1
        Dim b As Integer = 2
        Dim c As Integer = 3
        Dim d As Integer = 4
        Dim f As Integer = 5
        Dim g As Integer = 6
        ListBox1.Items.Add(10 ^ ((a - 1) / 6))
        ListBox1.Items.Add(10 ^ ((b - 1) / 6) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((c - 1) / 6) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((d - 1) / 6) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((f - 1) / 6) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((g - 1) / 6) * (TextBox3.Text / 10))
        CheckBox1.Enabled = False
    End Sub


    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        CheckBox1.Enabled = False
        CheckBox3.Enabled = False
        Dim a As Integer = 1
        Dim b As Integer = 2
        Dim c As Integer = 3
        Dim d As Integer = 4
        Dim f As Integer = 5
        Dim g As Integer = 6
        Dim h As Integer = 7
        Dim i As Integer = 8
        Dim j As Integer = 9
        Dim k As Integer = 10
        Dim l As Integer = 11
        Dim m As Integer = 12
        ListBox1.Items.Add(10 ^ ((a - 1) / 12))
        ListBox1.Items.Add(10 ^ ((b - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((c - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((d - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((f - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((g - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((h - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((i - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((j - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((k - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((l - 1) / 12) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((m - 1) / 12) * (TextBox3.Text / 10))
        CheckBox2.Enabled = False
    End Sub


    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        Dim a As Integer = 1
        Dim b As Integer = 2
        Dim c As Integer = 3
        Dim d As Integer = 4
        Dim f As Integer = 5
        Dim g As Integer = 6
        Dim h As Integer = 7
        Dim i As Integer = 8
        Dim j As Integer = 9
        Dim k As Integer = 10
        Dim l As Integer = 11
        Dim m As Integer = 12
        Dim a1 As Integer = 13
        Dim b1 As Integer = 14
        Dim c1 As Integer = 15
        Dim d1 As Integer = 16
        Dim f1 As Integer = 17
        Dim g1 As Integer = 18
        Dim h1 As Integer = 19
        Dim i1 As Integer = 20
        Dim j1 As Integer = 21
        Dim k1 As Integer = 22
        Dim l1 As Integer = 23
        Dim m1 As Integer = 24
        ListBox1.Items.Add(10 ^ ((a - 1) / 24))
        ListBox1.Items.Add(10 ^ ((b - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((c - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((d - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((f - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((g - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((h - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((i - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((j - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((k - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((l - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((m - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((a1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((b1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((c1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((d1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((f1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((g1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((h1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((i1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((j1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((k1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((l1 - 1) / 24) * (TextBox3.Text / 10))
        ListBox1.Items.Add(10 ^ ((m1 - 1) / 24) * (TextBox3.Text / 10))
        CheckBox3.Enabled = False
    End Sub
End Class
