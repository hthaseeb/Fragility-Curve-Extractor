Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim screenHeight As Integer = My.Computer.Screen.Bounds.Height
        Dim screenWidth As Integer = My.Computer.Screen.Bounds.Width
        Me.Height = screenHeight - 35
        Me.Width = screenWidth
        TabControl1.Height = screenHeight - 45
        TabControl1.Width = screenWidth - 10
        TabControl2.Height = screenHeight - 55
        TabControl2.Width = screenWidth - 10

        TextBox1.Text = 1.25
        TextBox2.Text = 0.2
        TextBox3.Text = 1.5
        TextBox4.Text = 0.3
        TextBox6.Text = 0
        TextBox7.Text = 5
        TextBox12.Text = 0.01

        TextBox57.Text = 2
        TextBox55.Text = 0.4
        TextBox53.Text = 0
        TextBox52.Text = 6
        TextBox51.Text = 0.001

        TextBox28.Text = 1
        TextBox26.Text = 0.4
        TextBox31.Text = 0
        TextBox30.Text = 7
        TextBox29.Text = 0.001

        TextBox38.Text = 1
        TextBox36.Text = 0.4
        TextBox34.Text = 0
        TextBox33.Text = 7
        TextBox32.Text = 0.001



        TextBox24.Text = 1.25
        TextBox23.Text = 0.2
        TextBox22.Text = 1.5
        TextBox21.Text = 0.3
        TextBox19.Text = 0
        TextBox18.Text = 5
        TextBox13.Text = 0.01

        TextBox50.Text = 4
        TextBox48.Text = 0
        TextBox49.Text = 0
        TextBox46.Text = 0
        TextBox45.Text = 5
        TextBox40.Text = 0.01

        TextBox42.Text = 5.5865
        TextBox44.Text = 0.739
        TextBox41.Text = 11.149
        TextBox43.Text = 0.912
        TextBox39.Text = 13.554
        TextBox37.Text = 1.36
        TextBox17.Text = 0.001

        TextBox24.Text = 1.25
        TextBox23.Text = 0.2
        TextBox22.Text = 1.5
        TextBox21.Text = 0.3
        TextBox19.Text = 0
        TextBox18.Text = 5
        TextBox13.Text = 0.01

    End Sub
   

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Cursor = Cursors.WaitCursor

        Dim x1 As Double
        Dim y1 As Double
        Dim x2 As Double
        Dim y2 As Double
        Dim min As Double
        Dim max As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p1 As Double
        Dim p2 As Double
        Dim p As Double
        Dim check As Integer
        Dim theta As Double
        Dim diff As Double
        Dim diffcheck As Double

        x1 = Nothing
        y1 = Nothing
        x2 = Nothing
        y2 = Nothing
        min = Nothing
        max = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p1 = Nothing
        p2 = Nothing
        p = Nothing
        theta = Nothing
        diff = Nothing
        diffcheck = 1000
        check = 0
        Me.Chart1.Series(0).Points.Clear()
        Me.Chart1.Series(1).Points.Clear()
        Me.DataGridView1.Rows.Clear()


        If Not IsNumeric(TextBox1.Text) Then
            TextBox1.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox2.Text) Then
            TextBox2.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox3.Text) Then
            TextBox3.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox4.Text) Then
            TextBox4.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox6.Text) Then
            TextBox6.Clear()
            Cursor = Cursors.Default
            MsgBox("Please enter numbers only.", vbInformation)
            Exit Sub
        ElseIf Not IsNumeric(TextBox7.Text) Then
            TextBox7.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox12.Text) Then
            TextBox12.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If




        x1 = TextBox1.Text
        y1 = TextBox2.Text
        x2 = TextBox3.Text
        y2 = TextBox4.Text
        min = TextBox6.Text
        max = TextBox7.Text
        interval = TextBox12.Text


        For m As Double = min To max Step interval
            For d As Double = 0 To 1 Step 0.001
                mu = Math.Log(m)


                p1 = MathNet.Numerics.Distributions.LogNormal.CDF(mu, d, x1)
                p2 = MathNet.Numerics.Distributions.LogNormal.CDF(mu, d, x2)


                diff = (y1 - p1) ^ 2 + (y2 - p2) ^ 2
                If diff < diffcheck Then
                    theta = m
                    sigma = d
                    diffcheck = diff
                End If

            Next
        Next



        TextBox8.Text = Math.Round(theta, 3)
        TextBox9.Text = Math.Round(sigma, 3)



        For x As Double = min To max Step interval

            p = MathNet.Numerics.Distributions.LogNormal.CDF(Math.Log(theta), sigma, x)

            Me.Chart1.Series(0).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
            Me.DataGridView1.Rows.Add(Math.Round(x, 3), Math.Round(p, 3))
        Next


        Me.Chart1.Series(1).Points.AddXY(x1, y1)
        Me.Chart1.Series(1).Points.AddXY(x2, y2)




        Cursor = Cursors.Default

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cursor = Cursors.WaitCursor
        Dim x1 As Double
        Dim y1 As Double
        Dim x2 As Double
        Dim y2 As Double
        Dim min As Double
        Dim max As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p1 As Double
        Dim p2 As Double
        Dim p As Double
        Dim check As Integer
        Dim theta As Double
        Dim diff As Double
        Dim diffcheck As Double

        x1 = Nothing
        y1 = Nothing
        x2 = Nothing
        y2 = Nothing
        min = Nothing
        max = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p1 = Nothing
        p2 = Nothing
        p = Nothing
        theta = Nothing
        diff = Nothing
        diffcheck = 1000
        check = 0
        Me.Chart2.Series(0).Points.Clear()
        Me.Chart2.Series(1).Points.Clear()
        Me.Chart2.Series(2).Points.Clear()



        Me.DataGridView2.Rows.Clear()

        If Not IsNumeric(TextBox24.Text) Then
            TextBox24.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox23.Text) Then
            TextBox23.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox22.Text) Then
            TextBox22.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox21.Text) Then
            TextBox21.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox19.Text) Then
            TextBox19.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox18.Text) Then
            TextBox18.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox13.Text) Then
            TextBox13.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If




        x1 = TextBox24.Text
        y1 = TextBox23.Text
        x2 = TextBox22.Text
        y2 = TextBox21.Text
        min = TextBox19.Text
        max = TextBox18.Text
        interval = TextBox13.Text


        For m As Double = min To max Step interval
            For d As Double = 0 To 1 Step 0.001
                mu = Math.Log(m)


                p1 = MathNet.Numerics.Distributions.LogNormal.PDF(mu, d, x1)
                p2 = MathNet.Numerics.Distributions.LogNormal.PDF(mu, d, x2)



                diff = (y1 - p1) ^ 2 + (y2 - p2) ^ 2
                If diff < diffcheck Then
                    theta = m
                    sigma = d
                    diffcheck = diff
                End If

            Next
        Next



        TextBox16.Text = Math.Round(theta, 3)
        TextBox15.Text = Math.Round(sigma, 3)

        For x As Double = min To max Step interval

            p = MathNet.Numerics.Distributions.LogNormal.CDF(Math.Log(theta), sigma, x)

            Me.Chart2.Series(0).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
            Me.DataGridView2.Rows.Add(Math.Round(x, 3), Math.Round(p, 3))

        Next
        For x As Double = min To max Step interval

            p = MathNet.Numerics.Distributions.LogNormal.PDF(Math.Log(theta), sigma, x)

            Me.Chart2.Series(1).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
        Next

        Me.Chart2.Series(2).Points.AddXY(x1, y1)
        Me.Chart2.Series(2).Points.AddXY(x2, y2)



        Cursor = Cursors.Default
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Cursor = Cursors.WaitCursor
        Dim mean As Double
        Dim variance As Double
        Dim min As Double
        Dim max As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p As Double
        Dim theta As Double
        Dim beta As Double

        min = Nothing
        max = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p = Nothing
        mean = Nothing
        variance = Nothing
        theta = Nothing
        beta = Nothing
        Me.Chart3.Series(0).Points.Clear()
        Me.DataGridView3.Rows.Clear()

        If Not IsNumeric(TextBox28.Text) Then
            TextBox28.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox26.Text) Then
            TextBox26.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox31.Text) Then
            TextBox31.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox30.Text) Then
            TextBox30.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox29.Text) Then
            TextBox29.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If

        mean = TextBox28.Text
        variance = TextBox26.Text
        min = TextBox31.Text
        max = TextBox30.Text
        interval = TextBox29.Text

        mu = Math.Log((mean ^ 2) / Math.Sqrt(variance + mean ^ 2))
        sigma = Math.Sqrt(Math.Log(variance / (mean ^ 2) + 1))
        beta = sigma
        theta = mean / Math.Sqrt(Math.Exp(beta ^ 2))


        TextBox11.Text = Math.Round(theta, 3)
        TextBox10.Text = Math.Round(beta, 3)

        For x As Double = min To max Step interval

            p = MathNet.Numerics.Distributions.LogNormal.CDF(mu, sigma, x)


            Me.Chart3.Series(0).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
            Me.DataGridView3.Rows.Add(Math.Round(x, 3), Math.Round(p, 3))

        Next


        Cursor = Cursors.Default
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Cursor = Cursors.WaitCursor
        Dim mean As Double
        Dim cov As Double
        Dim variance As Double
        Dim min As Double
        Dim max As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p As Double
        Dim theta As Double
        Dim beta As Double


        cov = Nothing
        min = Nothing
        max = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p = Nothing
        mean = Nothing
        variance = Nothing
        theta = Nothing
        beta = Nothing
        Me.Chart4.Series(0).Points.Clear()
        Me.DataGridView4.Rows.Clear()

        If Not IsNumeric(TextBox38.Text) Then
            TextBox38.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox36.Text) Then
            TextBox36.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox34.Text) Then
            TextBox34.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox33.Text) Then
            TextBox33.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox32.Text) Then
            TextBox32.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If

        mean = TextBox38.Text
        cov = TextBox36.Text
        min = TextBox34.Text
        max = TextBox33.Text
        interval = TextBox32.Text

        variance = (mean * cov) ^ 2
        mu = Math.Log((mean ^ 2) / Math.Sqrt(variance + mean ^ 2))
        sigma = Math.Sqrt(Math.Log(variance / (mean ^ 2) + 1))
        beta = sigma
        theta = mean / Math.Sqrt(Math.Exp(beta ^ 2))



        TextBox25.Text = Math.Round(theta, 3)
        TextBox14.Text = Math.Round(beta, 3)

        For x As Double = min To max Step interval

            p = MathNet.Numerics.Distributions.LogNormal.CDF(mu, sigma, x)

            Me.Chart4.Series(0).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
            Me.DataGridView4.Rows.Add(Math.Round(x, 3), Math.Round(p, 3))

        Next





        Cursor = Cursors.Default
    End Sub

    
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Cursor = Cursors.WaitCursor
        Dim x1 As Double
        Dim y1 As Double
        Dim x2 As Double
        Dim y2 As Double
        Dim min As Double
        Dim max As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p As Double
        Dim check As Integer
        Dim l As Integer
        Dim theta As Double
        Dim diff As Double
        Dim diffcheck As Double


        l = Nothing
        x1 = Nothing
        y1 = Nothing
        x2 = Nothing
        y2 = Nothing
        min = Nothing
        max = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p = Nothing
        theta = Nothing
        diffcheck = 1000


        check = 0
        Me.Chart5.Series(0).Points.Clear()
        Me.Chart5.Series(1).Points.Clear()
        Me.DataGridView5.Rows.Clear()


        If Not IsNumeric(TextBox40.Text) Then
            TextBox40.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox45.Text) Then
            TextBox45.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox46.Text) Then
            TextBox46.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox50.Text) Then
            TextBox50.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If


        l = TextBox50.Text
        Dim coor(l, 2) As Double
        Dim prob(l) As Double

        

        If l <> TextBox48.Lines.Count Or l <> TextBox49.Lines.Count Or TextBox48.Lines.Count <> TextBox49.Lines.Count Then
            MsgBox("Number of Coordinates is not Right", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If

        For x As Integer = 0 To l - 1
            coor(x, 1) = TextBox48.Lines(x)
            coor(x, 2) = TextBox49.Lines(x)

        Next



        min = TextBox46.Text
        max = TextBox45.Text
        interval = TextBox40.Text






        For m As Double = min To max Step interval
            For d As Double = 0 To 1 Step 0.001
                mu = Math.Log(m)

                For x As Integer = 0 To l - 1
                    prob(x) = MathNet.Numerics.Distributions.LogNormal.CDF(mu, d, coor(x, 1))
                Next

                diff = 0
                For x As Integer = 0 To l - 1

                    diff = diff + (prob(x) - coor(x, 2)) ^ 2

                Next


                If diff < diffcheck Then
                        theta = m
                        sigma = d
                        diffcheck = diff
                    End If

            Next

        Next






        TextBox35.Text = Math.Round(theta, 3)
        TextBox27.Text = Math.Round(sigma, 3)
        For x As Double = min To max Step interval

            p = MathNet.Numerics.Distributions.LogNormal.CDF(Math.Log(theta), sigma, x)

            Me.Chart5.Series(0).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
            Me.DataGridView5.Rows.Add(Math.Round(x, 3), Math.Round(p, 3))
        Next

        For x As Integer = 0 To l - 1

            Me.Chart5.Series(1).Points.AddXY(coor(x, 1), coor(x, 2))
        Next




        Cursor = Cursors.Default

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        

        Cursor = Cursors.WaitCursor
        Dim theta As Double
        Dim beta As Double
        Dim mean As Double
        Dim variance As Double
        Dim min As Double
        Dim max As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p As Double

        min = Nothing
        max = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p = Nothing
        mean = Nothing
        variance = Nothing
        theta = Nothing
        beta = Nothing
        Me.Chart6.Series(0).Points.Clear()
        Me.DataGridView6.Rows.Clear()

        If Not IsNumeric(TextBox57.Text) Then
            TextBox57.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox55.Text) Then
            TextBox55.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox53.Text) Then
            TextBox53.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox52.Text) Then
            TextBox52.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox51.Text) Then
            TextBox51.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If

        theta = TextBox57.Text
        beta = TextBox55.Text
        min = TextBox53.Text
        max = TextBox52.Text
        interval = TextBox51.Text

        mean = theta * Math.Sqrt(Math.Exp(beta ^ 2))
        variance = mean ^ 2 * (Math.Exp(beta ^ 2) - 1)

        For x As Double = min To max Step interval
            mu = Math.Log((mean ^ 2) / Math.Sqrt(variance + mean ^ 2))
            sigma = Math.Sqrt(Math.Log(variance / (mean ^ 2) + 1))
            p = MathNet.Numerics.Distributions.LogNormal.CDF(mu, sigma, x)

            
            Me.Chart6.Series(0).Points.AddXY(Math.Round(x, 3), Math.Round(p, 3))
            Me.DataGridView6.Rows.Add(Math.Round(x, 3), Math.Round(p, 3))

        Next


        Cursor = Cursors.Default
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Cursor = Cursors.WaitCursor
        ProgressBar1.Value = 0
        Me.Chart7.Series(0).Points.Clear()

        Dim sa16 As Double
        Dim s16 As Double
        Dim m16 As Double
        Dim sa50 As Double
        Dim s50 As Double
        Dim m50 As Double
        Dim sa84 As Double
        Dim s84 As Double
        Dim m84 As Double
        Dim interval As Double
        Dim mu As Double
        Dim sigma As Double
        Dim p1 As Double
        Dim p2 As Double
        Dim p3 As Double
        Dim p As Double
        Dim check As Integer
        Dim theta As Double
        Dim maxedp As Double
        Dim progress As Integer
        Dim diff As Double
        Dim diffcheck As Double


        sa16 = Nothing
        s16 = Nothing
        m16 = Nothing
        sa50 = Nothing
        s50 = Nothing
        m50 = Nothing
        sa84 = Nothing
        s84 = Nothing
        m84 = Nothing
        interval = Nothing
        mu = Nothing
        sigma = Nothing
        p1 = Nothing
        p2 = Nothing
        p3 = Nothing
        p = Nothing
        theta = Nothing
        maxedp = Nothing

        check = 0


        Me.DataGridView7.Rows.Clear()


        If Not IsNumeric(TextBox37.Text) Then
            TextBox37.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox39.Text) Then
            TextBox39.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox41.Text) Then
            TextBox41.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox42.Text) Then
            TextBox42.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        ElseIf Not IsNumeric(TextBox43.Text) Then
            TextBox43.Clear()
            Cursor = Cursors.Default
            MsgBox("Please enter numbers only.", vbInformation)
            Exit Sub
        ElseIf Not IsNumeric(TextBox44.Text) Then
            TextBox44.Clear()
            Cursor = Cursors.Default
            MsgBox("Please enter numbers only.", vbInformation)
            Exit Sub
        ElseIf Not IsNumeric(TextBox17.Text) Then
            TextBox17.Clear()
            MsgBox("Please enter numbers only.", vbInformation)
            Cursor = Cursors.Default
            Exit Sub
        End If


        s16 = TextBox42.Text
        sa16 = TextBox44.Text
        s50 = TextBox41.Text
        sa50 = TextBox43.Text
        s84 = TextBox39.Text
        sa84 = TextBox37.Text
        interval = TextBox17.Text
        m16 = sa16 / s16
        m50 = sa50 / s50
        m84 = sa84 / s84

        maxedp = 1.3 * Math.Max((Math.Max(m16, m50)), m84)
        progress = maxedp / interval
        ProgressBar1.Maximum = progress + 1
        progress = 0
        For edp As Double = 0 To maxedp Step interval


            mu = Nothing
            sa50 = Nothing
            sigma = Nothing
            diff = Nothing
            diffcheck = 1000


            If edp < m16 Then
                sa16 = edp * s16
            Else
                sa16 = s16 * m16
            End If

            If edp < m50 Then
                sa50 = edp * s50
            Else
                sa50 = s50 * m50
            End If

            If edp < m84 Then
                sa84 = edp * s84
            Else
                sa84 = s84 * m84
            End If

            mu = Math.Log(sa50)

            For d As Double = 0 To 1 Step 0.001



                p1 = MathNet.Numerics.Distributions.LogNormal.CDF(mu, d, sa16)
                p2 = MathNet.Numerics.Distributions.LogNormal.CDF(mu, d, sa84)
                p3 = MathNet.Numerics.Distributions.LogNormal.CDF(mu, d, sa50)

                diff = (0.16 - p1) ^ 2 + (0.84 - p2) ^ 2 + (0.5 - p3) ^ 2
                If diff < diffcheck Then
                    sigma = d
                    diffcheck = diff
                End If


            Next

            theta = sa50

            Me.Chart7.Series(0).Points.AddXY(edp, Math.Round(sigma, 3))
            Me.DataGridView7.Rows.Add(edp, sa50, Math.Round(theta, 3), mu, Math.Round(sigma, 3), diff, sa50)

            ProgressBar1.Value = progress

            progress = progress + 1
        Next
        ProgressBar1.Value = ProgressBar1.Maximum



        Cursor = Cursors.Default

    End Sub


End Class
