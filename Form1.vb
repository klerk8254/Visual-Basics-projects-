Imports MySql.Data.MySqlClient



Public Class Form1
    Dim con As MySqlConnection
    Dim cm As MySqlCommand
    Dim dt As New DataTable
    Dim reader As MySqlDataReader
    Dim bitmap As Bitmap



    Sub updatetable()
        con = New MySqlConnection("server = localhost; user=root; password=; database=project")
        con.Open()
        Dim query As String
        query = "select * from project.customer"
        cm = New MySqlCommand(query, con)
        reader = cm.ExecuteReader
        dt.Load(reader)
        reader.Close()
        con.Close()
        DataGridView1.DataSource = dt

    End Sub



    Private Function Cost_of_Items() As Double
        Dim sum As Double = 0
        Dim i As Integer = 0

        For i = 0 To DataGridView1.Rows.Count - 1
            sum = sum + Convert.ToDouble(DataGridView1.Rows(i).Cells(2).Value)
        Next i
        Return sum
    End Function



    Sub AddCost()
        Dim Tax, q As Double
        Tax = 3.5

        If DataGridView1.Rows.Count > 0 Then
            lblTax.Text = FormatCurrency(((Cost_of_Items() * Tax / 100).ToString("0.00")))
            lblSubTotal.Text = FormatCurrency(Cost_of_Items().ToString("0.00"))
            q = (Cost_of_Items() * Tax / 100)
            lblTotal.Text = FormatCurrency(q + Cost_of_Items().ToString("0.00"))
        End If
    End Sub

    Sub Change()
        Dim Tax, q, c As Double
        Tax = 3.5

        If DataGridView1.Rows.Count > 0 Then

            q = (Cost_of_Items() * Tax / 100) + Cost_of_Items()
            c = Val(lblCash.Text)
            lblChange.Text = FormatCurrency((c - q).ToString("0.00"))
        End If
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        lblChange.Text = "0"
        lblCash.Text = "0"
        lblSubTotal.Text = ""
        lblTax.Text = ""
        lblTotal.Text = ""
        CboPayment.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView1.Refresh()
    End Sub

    Private Sub NumbersOnly(sender As Object, e As EventArgs) Handles num1.Click, num6.Click, num5.Click, num4.Click, num9.Click, num8.Click, btnDot.Click, num0.Click, num3.Click, num2.Click, num7.Click, btnDot.Click
        Dim b As Button = sender

        If (lblCash.Text = "0") Then
            lblCash.Text = ""
            lblCash.Text = b.Text

        ElseIf (b.Text = ".") Then
            If (Not lblCash.Text.Contains(".")) Then
                lblCash.Text = lblCash.Text + b.Text

            End If
        Else
            lblCash.Text = lblCash.Text + b.Text
        End If

    End Sub



    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        lblCash.Text = "0"
        CboPayment.Text = ""
        lblChange.Text = ""
    End Sub

    'Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint
    '  CboPayment.Items.Add("Cash")
    ' CboPayment.Items.Add("Direct Debit ")
    ' CboPayment.Items.Add("Visa Card")
    ' CboPayment.Items.Add("Master Card")
    ' End Sub

    Private Sub btnPay_Click(sender As Object, e As EventArgs) Handles btnPay.Click
        If (CboPayment.Text = "Cash") Then
            Change()
        Else
            lblChange.Text = "0"
            lblCash.Text = "0"
        End If
    End Sub


    Private Sub btnRoveitem_Click(sender As Object, e As EventArgs) Handles btnRoveitem.Click

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            DataGridView1.Rows.Remove(row)
        Next
        AddCost()

        If (CboPayment.Text = "Cash") Then
            Change()
        Else
            lblChange.Text = ""
            lblCash.Text = ""
        End If

    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Dim height As Integer = DataGridView1.Height
        DataGridView1.Height = (DataGridView1.RowCount + 1) * DataGridView1.RowTemplate.Height
        bitmap = New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        DataGridView1.DrawToBitmap(bitmap, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
        PrintPreviewDialog1.ShowDialog()
        DataGridView1.Height = height

    End Sub

    Private Sub btnsilkGreek_Click(sender As Object, e As EventArgs) Handles btnsilkGreek.Click
        Dim CostOfItem As Double = 7.5
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "silkGreek" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("silkGreek", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnstonyfield_Click(sender As Object, e As EventArgs) Handles btnstonyfield.Click
        Dim CostOfItem As Double = 10.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "stonyfield" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("stonyfield", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnguavaberry_Click(sender As Object, e As EventArgs) Handles btnguavaberry.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "guavaberry" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
            End If
        Next

        DataGridView1.Rows.Add("guavaberry", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnpearcrystal_Click(sender As Object, e As EventArgs) Handles btnpearcrystal.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "pearcystal" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("pearcystal", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnmevgal_Click(sender As Object, e As EventArgs) Handles btnmevgal.Click
        Dim CostOfItem As Double = 9.5
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "mevgal" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("mevgal", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnpanamilk_Click(sender As Object, e As EventArgs) Handles btnpanamilk.Click
        Dim CostOfItem As Double = 15.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Pana Milk" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Pana Milk", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnlivoat_Click(sender As Object, e As EventArgs) Handles btnLivoat.Click
        Dim CostOfItem As Double = 25.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Livoat" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Livoat", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnovaltine_Click(sender As Object, e As EventArgs) Handles btnovaltine.Click
        Dim CostOfItem As Double = 18.09
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "ovaltine" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("ovaltine", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btndannon_Click(sender As Object, e As EventArgs) Handles btndannon.Click
        Dim CostOfItem As Double = 8.05
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Dannon" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Dannon", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawImage(bitmap, 10, 10)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        updateTable()






    End Sub

    Private Sub CboPayment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboPayment.SelectedIndexChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub lblCash_Click(sender As Object, e As EventArgs) Handles lblCash.Click

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnmango_Click(sender As Object, e As EventArgs) Handles btnmango.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Mango" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Mango", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnvanillaberry_Click(sender As Object, e As EventArgs) Handles btnvanillaberry.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "VanillaBerry" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("VanillaBerry", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnstrawberry_Click(sender As Object, e As EventArgs) Handles btnstarwberry.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "StrawBerry" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("StrawBerry", "1", CostOfItem)
        AddCost()

    End Sub

    Private Sub btnlimonid_Click(sender As Object, e As EventArgs) Handles btnlimonid.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Limonid" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Limonid", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnmelonide_Click(sender As Object, e As EventArgs) Handles btnmelonide.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Melonide" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Melonide", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btndonsimon_Click(sender As Object, e As EventArgs) Handles btndonsimon.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Don Simon" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Don Simon", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub btnneshua_Click(sender As Object, e As EventArgs) Handles btnneshua.Click
        Dim CostOfItem As Double = 12.0
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells(0).Value = "Neshua" Then
                row.Cells(1).Value = Double.Parse(row.Cells(1).Value + 1)
                row.Cells(2).Value = Double.Parse(row.Cells(1).Value) * CostOfItem
                Exit Sub
            End If
        Next

        DataGridView1.Rows.Add("Neshua", "1", CostOfItem)
        AddCost()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label0_Click(sender As Object, e As EventArgs) Handles Label0.Click

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub PrintPreviewDialog1_Load(sender As Object, e As EventArgs) Handles PrintPreviewDialog1.Load

    End Sub
End Class
