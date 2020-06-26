Module mod1
    Public f1 As Form1
    Public f2 As Form2
    Public f3 As Form3
    Public f4 As Form4
    Public open_w As Boolean = False
    Public p, p0 As Double : Public t1, t2, t3, t4, t As Integer
End Module
Public Class Form1
    Dim A, V, Te1, Te2, L1, L, L2, d1, d2, S1, S2, q0, al, k0, qp0, qp1, k1, k2, k3, k4, h, mi, ma As Decimal
    Dim num_r, num_l, no As Integer '1
    Dim pre() As Decimal '1
    Dim tim() As Integer '1
    Dim b As New Drawing.Bitmap(800, 640)
    Dim pe, te As Double
    Dim nume As Integer = 1
    Dim exl As Object
    Dim n As Integer
    Dim show_r As Boolean = True
    Dim dr_pn As Pen
    Dim num As Integer = 1
    Dim el1, el2, el3, el4, t_den As Integer
    Dim ti, tp, i As Integer
    Dim k01 As Decimal
    Structure rememder '1
        Dim p_r() As Decimal : Dim t_r() As Integer : Dim A_r As Decimal : Dim V_r As Decimal : Dim Te1_r As Decimal : Dim Te2_r As Decimal : Dim L1_r As Decimal : Dim l_r As Decimal
        Dim L2_r As Decimal : Dim d1_r As Decimal : Dim d2_r As Decimal : Dim S1_r As Decimal : Dim S2_r As Decimal : Dim q0_r As Decimal : Dim al_r As Decimal : Dim k0_r As Decimal
        Dim qp0_r As Decimal : Dim qp1_r As Decimal : Dim k1_r As Decimal : Dim k2_r As Decimal : Dim k3_r As Decimal : Dim k4_r As Decimal : Dim h_r As Decimal
        Dim t1_r As Decimal : Dim t2_r As Decimal : Dim t3_r As Decimal : Dim t4_r As Decimal : Dim p0_r As Decimal
        Dim numb As Integer : Dim k01_r As Decimal
    End Structure '1
    Dim rememd_pr(10) As rememder '1
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        num_l = ComboBox2.SelectedIndex
        Button5.Enabled = True
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            load_n()
        Catch ex As Exception
            MsgBox("данные пусты, загрузите их")
        End Try
    End Sub
    Sub load_n()
        Try
            TextBox1.Text = rememd_pr(num_l).t1_r : TextBox2.Text = rememd_pr(num_l).t2_r : TextBox3.Text = rememd_pr(num_l).t3_r
            TextBox4.Text = rememd_pr(num_l).t4_r : TextBox5.Text = rememd_pr(num_l).S1_r : TextBox6.Text = rememd_pr(num_l).d1_r : TextBox7.Text = rememd_pr(num_l).L1_r
            TextBox8.Text = rememd_pr(num_l).L2_r : TextBox9.Text = rememd_pr(num_l).d2_r : TextBox10.Text = rememd_pr(num_l).S2_r
            TextBox11.Text = rememd_pr(num_l).A_r : TextBox12.Text = rememd_pr(num_l).V_r : TextBox13.Text = rememd_pr(num_l).p0_r : TextBox14.Text = rememd_pr(num_l).Te1_r
            TextBox15.Text = rememd_pr(num_l).Te2_r : TextBox16.Text = rememd_pr(num_l).q0_r : TextBox17.Text = rememd_pr(num_l).k0_r
            TextBox18.Text = rememd_pr(num_l).al_r : TextBox20.Text = rememd_pr(num_l).qp1_r : TextBox21.Text = rememd_pr(num_l).k01_r : TextBox19.Text = rememd_pr(num_l).qp0_r
        Catch ex As Exception
            MsgBox("неверный ввод чисел")
        End Try
    End Sub
    Sub remem_num()
        rememd_pr(num_r).A_r = A : rememd_pr(num_r).V_r = V : rememd_pr(num_r).Te1_r = Te1 : rememd_pr(num_r).Te2_r = Te2 : rememd_pr(num_r).L1_r = L1 : rememd_pr(num_r).l_r = L
        rememd_pr(num_r).L2_r = L2 : rememd_pr(num_r).d1_r = d1 : rememd_pr(num_r).d2_r = d2 : rememd_pr(num_r).S1_r = S1 : rememd_pr(num_r).S2_r = S2 : rememd_pr(num_r).q0_r = q0
        rememd_pr(num_r).al_r = al : rememd_pr(num_r).k0_r = k0 : rememd_pr(num_r).qp0_r = qp0 : rememd_pr(num_r).qp1_r = qp1 : rememd_pr(num_r).k1_r = k1 : rememd_pr(num_r).k2_r = k2
        rememd_pr(num_r).k3_r = k3 : rememd_pr(num_r).k4_r = k4 : rememd_pr(num_r).h_r = h : rememd_pr(num_r).t1_r = t1 : rememd_pr(num_r).t2_r = t2 - t1 : rememd_pr(num_r).t3_r = t3 - t2
        rememd_pr(num_r).t4_r = t4 - t3 : rememd_pr(num_r).p0_r = p0 : rememd_pr(num_r).k01_r = k01
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            exc_e()
            If exl.visible = False Then exl.visible = True
            exl.range("A1").value = "этап откачки"
            exl.range("B1").value = "время, мин."
            exl.range("C1").value = "давление, Па."
            wr_e()
            nume += el1 + el2 + el3 + el4 + 8
            num += 1
        Catch ex As Exception
            MsgBox("ошибка, произведите расчеты")
        End Try
    End Sub
    Sub exc_e()
        el1 = TextBox23.Text : el2 = TextBox24.Text : el3 = TextBox25.Text : el4 = TextBox26.Text
    End Sub
    Sub wr_e()
        Dim i, pu, k As Integer
        exl.range("A" & nume + 1 + k).value = "форврак. откачка"
        exl.range("B" & nume + 1 + k).value = Format(0 / 60, "0.##")
        exl.range("c" & nume + 1 + k).value = Format(p0, "0.####E+0")
        k += 1
        If t1 Mod el1 <> 0 Then
            pu = t1 Mod el1
        End If
        For i = 1 To t1 - pu Step t1 / el1
            exl.range("B" & nume + 1 + k).value = Format(tim(i) / 60, "0.##")
            exl.range("c" & nume + 1 + k).value = Format(pre(i), "0.####E+0")
            k += 1
        Next i
        exl.range("A" & nume + 1 + k).value = "вакуумная откачка"
        If (t2 - t1) Mod el2 <> 0 Then
            pu = (t2 - t1) Mod el2
        End If
        For i = t1 To t2 - pu Step ((t2 - t1) - pu) / el2
            exl.range("B" & nume + 1 + k).value = Format(tim(i) / 60, "0.##")
            exl.range("c" & nume + 1 + k).value = Format(pre(i), "0.####E+0")
            k += 1
        Next i
        exl.range("A" & nume + 1 + k).value = "вак. откач. с/нагр."
        If (t3 - t2) Mod el3 <> 0 Then
            pu = (t3 - t2) Mod el3
        End If
        For i = t2 To t3 - pu Step ((t3 - t2) - pu) / el3
            exl.range("B" & nume + 1 + k).value = Format(tim(i) / 60, "0.##")
            exl.range("c" & nume + 1 + k).value = Format(pre(i), "0.####E+0")
            k += 1
        Next i
        exl.range("A" & nume + 1 + k).value = "вак. откач. б/нагр."
        If (t4 - t3) Mod el4 <> 0 Then
            pu = (t4 - t3) Mod el4
        End If
        For i = t3 To (t4 - pu) - ((t4 - t3) - pu) / el4 Step ((t4 - t3) - pu) / el4
            exl.range("B" & nume + 1 + k).value = Format(tim(i) / 60, "0.##")
            exl.range("c" & nume + 1 + k).value = Format(pre(i), "0.####E+0")
            k += 1
        Next i
        exl.range("B" & nume + 1 + k).value = Format(tim(t4) / 60, "0.##")
        exl.range("c" & nume + 1 + k).value = Format(pre(t4), "0.####E+0")
        k += 2
        exl.range("a" & nume + 1 + k).value = "сохранение "
        exl.range("b" & nume + 1 + k).value = " N "
        exl.range("c" & nume + 1 + k).value = num
    End Sub
    Sub remem_q()
        ReDim rememd_pr(num_r).p_r(t4)
        ReDim rememd_pr(num_r).t_r(t4)
        For i = 0 To t4
            rememd_pr(num_r).p_r(i) = pre(i)
            rememd_pr(num_r).t_r(i) = tim(i)
        Next i
        rememd_pr(num_r).numb = t4
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim wri As New System.IO.StreamWriter("exp/inf.txt")
            wri.WriteLine("Параметры системы")
            wri.Write("Время откачки форвакуумным насосом:                    ") : wri.WriteLine(t1)
            wri.Write("Время откачки вакуумным насосм до вкл. нагрева камеры: ") : wri.WriteLine(t2)
            wri.Write("Время нагрева камеры:                                  ") : wri.WriteLine(t3)
            wri.Write("Время откачки после нагрева:                           ") : wri.WriteLine(t4)
            wri.WriteLine("Форвакуумный насос")
            wri.Write("Скорость откачки:                                      ") : wri.WriteLine(S1)
            wri.Write("Диаметр соединительного канала:                        ") : wri.WriteLine(d1)
            wri.Write("Длина соединительного канала:                          ") : wri.WriteLine(L1)
            wri.WriteLine("Вакуумный насос")
            wri.Write("Скорость откачки:                                      ") : wri.WriteLine(S2)
            wri.Write("Диаметр соединительного канала:                        ") : wri.WriteLine(d2)
            wri.Write("Длина соединительного канала:                          ") : wri.WriteLine(L2) '10
            wri.Write("Площадь поеверхности:                                  ") : wri.WriteLine(A)
            wri.Write("Объем:                                                 ") : wri.WriteLine(V)
            wri.Write("Начальное давление в системе:                          ") : wri.WriteLine(p0)
            wri.Write("Температура до нагрева:                                ") : wri.WriteLine(Te1)
            wri.Write("Температура нагрева                                    ") : wri.WriteLine(Te2)
            wri.Write("Удальное газовыделение стенок:                         ") : wri.WriteLine(q0)
            wri.Write("Коэф. скорости газовыделения:                          ") : wri.WriteLine(k0)
            wri.Write("Постоянный коэф.:                                      ") : wri.WriteLine(al)
            wri.Write("Поток до нагрева:                                      ") : wri.WriteLine(qp0)
            wri.Write("Поток после нагрева:                                   ") : wri.WriteLine(qp1)
            wri.WriteLine("Откачка форвакуумным насосом")
            wri.Write("давление ") : wri.Write((Format(p0, "0.###E+0"))) : wri.Write(" Па.,   время ") : wri.Write(0) : wri.WriteLine(" c.")
            For i = 1 To t4
                If i = t1 Then
                    wri.WriteLine("Откачка вакуумным насосом до вкл. нагрева камеры")
                ElseIf i = t2 Then
                    wri.WriteLine("Откачка вакуумным насосом во время нагрева камеры")
                ElseIf i = t3 Then
                    wri.WriteLine("Откачка вакуумным насосом при постоянной темпиратуре")
                End If
                wri.Write("давление ") : wri.Write((Format(pre(i), "0.###E+0"))) : wri.Write(" Па.,   время ") : wri.Write(tim(i)) : wri.WriteLine(" c.")
            Next i
            wri.Write("Полное время откачки:                              ") : wri.Write(t4) : wri.WriteLine(" с.")
            wri.Write("Минимальное давление:                              ") : wri.Write(pre(t4)) : wri.WriteLine(" Па.")
            wri.Close()
        Catch ex As Exception
            MsgBox("Необходимо провести эксперимент!")
        End Try
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            SaveFileDialog1.Filter = "Текстовые фаилы|*.txt"
            If SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then Exit Sub
            Dim fil As String = SaveFileDialog1.FileName
            Dim wr As New System.IO.StreamWriter(fil)
            wr.WriteLine(Format(p0, "0.###E+0") & vbTab & 0)
            For i = 1 To t4
                wr.WriteLine(Format(pre(i), "0.##E+0") & vbTab & tim(i))
            Next i
            wr.Close()
        Catch ex As Exception
            MsgBox("Ошибка: возможно эксперимент еще не проведен")
        End Try
    End Sub
    Private Sub СохранениеДанныхToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СохранениеДанныхToolStripMenuItem.Click
        f4 = New Form4
        f4.Button1.Focus()
        f4.Show()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        num_r = ComboBox1.SelectedIndex
        Button4.Enabled = True
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            remem_num()
            remem_q()
            FileOpen(1, "pres.save", OpenMode.Binary)
            FilePut(1, rememd_pr(num_r))
            FileClose(1)
        Catch ex As Exception
            MsgBox("Возможно эксперимент проведен с ошибкой, повторите эксперимент")
        End Try
    End Sub
    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        f3 = New Form3
        f3.Button1.Focus()
        f3.Show()
    End Sub
    Public Function g() As Graphics
        g = Graphics.FromImage(b)
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        f1 = Me

        TextBox1.Text = 650 : TextBox2.Text = 1000 : TextBox3.Text = 3000
        TextBox4.Text = 2500 : TextBox5.Text = 0.055 : TextBox6.Text = 25 * 10 ^ (-3) : TextBox7.Text = 0.5
        TextBox8.Text = 0.5 : TextBox9.Text = 40 * 10 ^ (-3) : TextBox10.Text = 0.31
        TextBox11.Text = 1.5 : TextBox12.Text = 1.5 : TextBox13.Text = 10 ^ 5 : TextBox14.Text = 293
        TextBox15.Text = 893 : TextBox16.Text = 1 * 10 ^ (-6) : TextBox17.Text = 1 * 10 ^ (-4)
        TextBox18.Text = -2 * 10 ^ 3 : TextBox20.Text = 10 ^ (-13) : TextBox19.Text = 10 ^ (-10)
        TextBox23.Text = 10 : TextBox24.Text = 20 : TextBox25.Text = 80 : TextBox26.Text = 10
        TextBox21.Text = 10 ^ (-10)
        TextBox27.Text = "информация эксперимента"

        If IO.Directory.Exists("exp") Then
        Else
            IO.Directory.CreateDirectory("exp")
        End If
        exl = CreateObject("excel.application")
        exl.workbooks.add()

        Button4.Enabled = False
        Button5.Enabled = False
    End Sub
    Sub Enter_w()
        Try
            t1 = TextBox1.Text : t2 = TextBox2.Text + t1 : t3 = TextBox3.Text + t2 : t4 = TextBox4.Text + t3 : S1 = TextBox5.Text : d1 = TextBox6.Text
            S2 = TextBox10.Text : d2 = TextBox9.Text : L2 = TextBox8.Text : L1 = TextBox7.Text : A = TextBox11.Text
            V = TextBox12.Text : p0 = TextBox13.Text : Te1 = TextBox14.Text : Te2 = TextBox15.Text : q0 = TextBox16.Text : k0 = TextBox17.Text
            al = TextBox18.Text : qp1 = TextBox20.Text : qp0 = TextBox19.Text : k01 = TextBox21.Text
            el1 = TextBox23.Text : el2 = TextBox24.Text : el3 = TextBox25.Text : el4 = TextBox26.Text
            TextBox27.Text = ""
        Catch ex As Exception
            MsgBox("вводить только числа")
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If open_w = True Then
            g.Clear(Color.Transparent)
            f2.Close()
            open_w = False
        End If
        If open_w = False Then
            f2 = New Form2
            Me.AddOwnedForm(f2)
            f2.Show()
            f2.PictureBox1.Image = b
            open_w = True
        End If
        Enter_w()
        Choose_sm()
        show_r = True
        Try
            fi()
            dro()
            Drawing_arc(g)
        Catch ex As Exception
            g.Clear(Color.Transparent)
            f2.Close()
            open_w = False
        End Try
    End Sub
    Sub Drawing_arc(ByVal g As Graphics) ' строим оси для графика
        Dim ft As New Font("arial", 10)
        Dim pu As Integer
        g.DrawLine(Pens.Black, 20, 20, 20, 620)
        g.DrawLine(Pens.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(no)))), 780, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(no)))))
        g.DrawString("t, мин.", ft, Brushes.Black, 740, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(no)))))
        g.DrawString("Lg(p)", ft, Brushes.Black, 20, 0)

        If t4 Mod 10 <> 0 Then
            pu = t4 Mod 10
        End If
        g.DrawString(0, ft, Brushes.Black, 25, CSng(pe * (Math.Log(p0) - Math.Log(pre(no)))))
        For i = 1 To 9
            g.DrawString(Format((i * (t4 - pu) / 10) / 60, "0.#"), ft, Brushes.Black, 20 + (760 / t4) * (i * t4 / 10), CSng(pe * (Math.Log(p0) - Math.Log(pre(no)))))
            g.DrawLine(Pens.Black, CSng(20 + (760 / t4) * (i * t4 / 10)), CSng(15 + pe * (Math.Log(p0) - Math.Log(pre(no)))), CSng(20 + (760 / t4) * (i * t4 / 10)), CSng(25 + pe * (Math.Log(p0) - Math.Log(pre(no)))))
        Next i

        g.DrawString(Format(t4 / 60, "0.#"), ft, Brushes.Black, 750, CSng(pe * (Math.Log(p0) - Math.Log(pre(no)))))
        g.DrawLine(Pens.Black, CSng(780), CSng(15 + pe * (Math.Log(p0) - Math.Log(pre(no)))), CSng(780), CSng(25 + pe * (Math.Log(p0) - Math.Log(pre(no)))))

        If pre.Length Mod 10 <> 0 Then
            pu = pre.Length Mod 10
        End If

        g.DrawString(Format(p0, "0.##E+0"), ft, Brushes.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(p0))))
        g.DrawLine(Pens.Black, 15, CSng(20 + pe * (Math.Log(p0) - Math.Log(p0))), 25, CSng(20 + pe * (Math.Log(p0) - Math.Log(p0))))

        If ma <> 0 Then
            g.DrawString(Format(ma, "0.##E+0"), ft, Brushes.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(ma))))
            g.DrawLine(Pens.Black, 15, CSng(20 + pe * (Math.Log(p0) - Math.Log(ma))), 25, CSng(20 + pe * (Math.Log(p0) - Math.Log(ma))))
        End If
        g.DrawString(Format(pre(t1), "0.##E+0"), ft, Brushes.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(t1)))))
        g.DrawLine(Pens.Black, 15, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(t1)))), 25, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(t1)))))

        g.DrawString(Format(pre(t2), "0.##E+0"), ft, Brushes.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(t2)))))
        g.DrawLine(Pens.Black, 15, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(t2)))), 25, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(t2)))))

        For i = t2 To t3 Step (t3 - t2) / 2
            g.DrawString(Format(pre(i), "0.##E+0"), ft, Brushes.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))))
            g.DrawLine(Pens.Black, 15, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))), 25, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))))
        Next i

        For i = t3 To t4 Step (t4 - t3) / 2
            g.DrawString(Format(pre(i), "0.##E+0"), ft, Brushes.Black, 20, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))))
            g.DrawLine(Pens.Black, 15, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))), 25, CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))))
        Next i
    End Sub
    Sub dro()
        For g1 = 1 To t4
            If (Math.Log(pre(g1)) < 0.1 And Math.Log(pre(g1)) > -0.1) Then
                no = g1
            End If
        Next
        pe = 600 / (Math.Log(p0) - Math.Log(mi))
        te = 760 / t4
        For i = 1 To t4 - 1
            g.DrawLine(Pens.Red, CSng(20 + te * tim(i)), CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i)))), CSng(20 + te * tim(i + 1)), CSng(20 + pe * (Math.Log(p0) - Math.Log(pre(i + 1)))))
        Next i
    End Sub
    Sub fi()
        mi = pre(1)
        For i = 1 To t4 - 1
            If pre(i) < mi Then
                mi = pre(i)
            End If
        Next
    End Sub
    Sub Choose_sm()
        ReDim pre(t4) '1
        ReDim tim(t4) '1
        Dim wr As New System.IO.StreamWriter("exp\exp.txt")
        p = p0
        wr.Write("давление ") : wr.Write((Format(p0, "0.##E+0"))) : wr.Write(" Па.,   время ") : wr.Write(0) : wr.WriteLine(" c.")
        For n = 1 To 4
            If p > 10 And n <> 1 Then
                MsgBox("Минимальное давление для включения вакуумного насоса не достигнуто")
                show_r = False
                Exit For
            Else
                Select Case n
                    Case 1
                        ti = 1 : tp = t1 : i = 1
                    Case 2
                        ti = t1 : tp = t2 : i = 2
                    Case 3
                        ti = t2 : tp = t3 : i = 2
                    Case 4
                        ti = t3 : tp = t4 : i = 2
                End Select
                Try
                    For t = ti To tp
                        h = (tp - ti) / (tp - ti)
                        p += Ks(i, t, p)
                        If n = 1 And p < 10 Then
                            TextBox27.Text = "Предельное давление для включения вакуумного насоса достигнуто на " & t & " секунде. Временной промежуток автоматически скорректирован."
                            reme()
                            TextBox1.Text = t1
                            Exit For
                        Else
                            pre(t) = p
                            tim(t) = t
                        End If
                        wr.Write("давление ") : wr.Write((Format(p, "0.###E+0"))) : wr.Write(" Па.,   время ") : wr.Write(t) : wr.WriteLine(" c.") '''''
                    Next t
                Catch ex As Exception
                End Try
            End If
        Next n
        wr.Close()
        If show_r Then
            TextBox27.Text = TextBox27.Text & vbCrLf _
                & "Время откачки составило " & Format(((t1 + t2 + t3 + t4) / 3600), "0.00") & " ч. " _
            & vbCrLf & "Предельное давление " & Format(p, "0.###E+0") & " Па."
        Else
            TextBox27.Text = TextBox27.Text & vbCrLf & "Время откачки составило " & Format((t1 / 3600), "0.00") & " ч." & vbCrLf & "Предельное давление " & Format(p, "0.###E+0") & " Па."
        End If
    End Sub
    Sub reme()
        Dim tet(t), pep(t) As Double
        t1 = t : t2 = TextBox2.Text + t1 : t3 = TextBox3.Text + t2 : t4 = TextBox4.Text + t3
        For i = 1 To t
            tet(i) = tim(i)
            pep(i) = pre(i)
        Next
        ReDim pre(t4) '1
        ReDim tim(t4) '1
        For i = 1 To t
            pre(i) = pep(i)
            tim(i) = tet(i)
        Next
    End Sub

    Function F(ByVal ch As Integer, ByVal t As Single, ByVal p As Decimal) As Decimal
        Dim pr1, pr2, pr3 As Decimal
        Dim xer As Decimal = k0 + (k01 - k0) * (t - t2) / (t3 - t2)
        pr1 = (Math.E) ^ (al * ((1 / Te1) - 1 / 293))
        pr2 = (Math.E) ^ (al * ((1 / Te2) - 1 / 293))
        pr3 = (Math.E) ^ (al * ((1 / (Te1 + (Te2 - Te1) / (t3 - t2) * (t - t2))) ^ 2 - (1 / 293)))
        If n = 1 Then
            F = ((-1) * Sef(ch, p) * p + (A * q0 * k0 * pr1 * ((Math.E) ^ (-k0 * t * pr1)) + qp0 * (p0 - p))) / V
        ElseIf n = 2 Then
            F = ((-1) * Sef(ch, p) * p + (A * q0 * k0 * pr1 * ((Math.E) ^ (-k0 * t * pr1)) + qp0 * (p0 - p))) / V
        ElseIf n = 3 Then
            F = ((-1) * Sef(ch, p) * p + (A * q0 * xer * pr3 *
                (Math.E) ^ (-xer * t * ((Math.E) ^ (al * ((1 / ((Te1 + (Te2 - Te1) * (t - t2)) / (t3 - t2))) ^ -2 - 1 / 293)))) + (qp1 - qp0) * (t - t2) / (t3 - t2)) * (p0 - p)) / V
        ElseIf n = 4 Then
            F = ((-1) * Sef(ch, p) * p + (A * q0 * k01 * pr2 * ((Math.E) ^ (-k01 * t * pr2)) + qp1 * (p0 - p))) / V
        End If
        Return F
    End Function
    Function Sef(ByVal ch As Integer, ByVal p As Decimal) As Decimal ' Sэф 
        Dim S, d As Double
        If ch = 1 Then
            S = S1 : d = d1 : L = L1
        ElseIf ch = 2 Then
            S = S2 : d = d2 : L = L2
        End If
        Sef = (S * (((1.36 * 10 ^ 3) * p * (d ^ 4) / L) + ((121 * d ^ 3) / L)) * Koef1(ch, p)) /
            (S + (((1.36 * 10 ^ 3) * p * (d ^ 4) / L) + ((121 * d ^ 3) / L) * Koef1(ch, p)))
    End Function
    Function Koef1(ByVal ch As Integer, ByVal p As Decimal) As Decimal
        Dim d As Double
        If ch = 1 Then
            d = d1
        ElseIf ch = 2 Then
            d = d2
        End If
        Koef1 = ((1 + 1.9 * (10 ^ 4) * d * p) / ((1 + 2.35 * (10 ^ 4) * d * p)))
    End Function
    Function Ks(ByVal i As Integer, ByVal t As Integer, ByVal p As Decimal) As Decimal ' рунге кутт
        k1 = F(i, t, p)
        k2 = F(i, t + h / 2, p + (h * k1) / 2)
        k3 = F(i, t + h / 2, p + (h * k2) / 2)
        k4 = F(i, t + h, p + h * k3)
        Ks = h * (k1 + 2 * k2 + 2 * k3 + k4) / 6
    End Function
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Application.Exit()
    End Sub
End Class
