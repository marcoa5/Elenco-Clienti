Imports Microsoft.SharePoint.Client
Module Module1
    Sub Main()
Inizio:
        Try
            SP()
        Catch ex As ClientRequestException
            If MsgBox($"Verifica di essere connesso alla rete aziendale{vbCr}{vbCr}Messaggio di Errore: {ex.Message}", vbRetryCancel, "Errore") = vbRetry Then
                GoTo Inizio
            Else
                Form1.Close()
            End If
        Catch ex As Exception
            MsgBox($"Errore: {ex.Message}", vbOK, "Errore")
            Form1.Close()
        End Try


    End Sub

    Dim Cliente(3000, 3)

    Public Property Cliente1 As Object(,)
        Get
            Return Cliente
        End Get
        Set(value As Object(,))
            Cliente = value
        End Set
    End Property

    Sub SP()
        Dim Path As String = "https://home.intranet.epiroc.com/sites/cc/iyc/"
        Dim I As Integer
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("Customers")
        Dim query As CamlQuery = CamlQuery.CreateAllItemsQuery(10000)
        Dim items As ListItemCollection = testList.GetItems(query)
        context.Load(items)
        context.ExecuteQuery()
        On Error Resume Next
        I = 0
        Form1.Show()

        Form1.ProgressBar1.Maximum = items.Count

        For Each listItem As ListItem In items
            Form1.ProgressBar1.Value = I

            Cliente(I, 0) = listItem("Title")
            Cliente(I, 1) = listItem("CustomerName")
            Cliente(I, 2) = listItem("Address_x0020_Line_x0020_1")
            Cliente(I, 3) = listItem("Address_x0020_Line_x0020_2")
            Form1.ComboBox1.Items.Add(listItem("CustomerName"))
            I += 1
        Next

        With Form1
            .ProgressBar1.Visible = False
            .ComboBox1.Visible = True
            .TextBox1.Visible = True
            .TextBox2.Visible = True
            .TextBox3.Visible = True
            .CheckBox1.Visible = True
            .CheckBox2.Visible = True
            .CheckBox3.Visible = True
            .CheckBox4.Visible = True
            .CheckBox1.Checked = True
            .CheckBox2.Checked = True
            .CheckBox3.Checked = True
            .ComboBox1.Select()
            .CheckBox1.Enabled = False
            .CheckBox2.Enabled = False
            .CheckBox3.Enabled = False
            .CheckBox4.Enabled = False
            .Button1.Enabled = False
            .ComboBox1.SelectedIndex = 1
        End With
    End Sub

    Sub Cambio()
        On Error GoTo eRRO
        With Form1
            If .ComboBox1.Text = "" Then
                .TextBox1.Text = ""
                .TextBox2.Text = ""
                .TextBox3.Text = ""
                .CheckBox1.Enabled = False
                .CheckBox2.Enabled = False
                .CheckBox3.Enabled = False
                .CheckBox4.Enabled = False
                .Button1.Enabled = False
            Else
                .TextBox1.Text = Cliente(Form1.ComboBox1.SelectedIndex, 2)
                Dim I1, Indirizzo As String
                Indirizzo = Cliente(Form1.ComboBox1.SelectedIndex, 3)

                If Mid(Indirizzo, Len(Indirizzo) - 2, 1) = " " Then
                    I1 = Right(Indirizzo, 2)
                    Indirizzo = Trim(Left(Indirizzo, Len(Indirizzo) - 2))
                    Indirizzo = Indirizzo & " (" & I1 & ")"
                End If
                .TextBox2.Text = Indirizzo
                .TextBox3.Text = Cliente(Form1.ComboBox1.SelectedIndex, 0)
                .CheckBox1.Enabled = True
                .CheckBox2.Enabled = True
                .CheckBox3.Enabled = True
                .CheckBox4.Enabled = True
                .Button1.Enabled = True
            End If
        End With
        Exit Sub
ErrO:
    End Sub


    Sub Copia()
        Dim Stringa As String = ""
        With Form1
            If .CheckBox1.Checked = True Then Stringa = .ComboBox1.Text & vbCrLf
            If .CheckBox2.Checked = True Then Stringa = Stringa & .TextBox1.Text & vbCrLf
            If .CheckBox3.Checked = True Then Stringa = Stringa & .TextBox2.Text & vbCrLf
            If .CheckBox4.Checked = True Then Stringa = Stringa & "Codice Cliente: " & .TextBox3.Text & vbCrLf

            Stringa = Left(Stringa, Len(Stringa) - 1)

            Clipboard.SetText(Stringa)
            .Close()
        End With

    End Sub
End Module
