Public Class Welding
    Dim mt, tc, vsh, ccr, ct, at, hof, dfcl, c, av, ts, p, hte, pt, Hnet As Double
    Dim tik, tin, TowTik, TowTin, ResultPerTemp As Double
    Dim x, y1, y2, y, ResultST As Double
    Dim z, x1, w, r, ResultWoH As Double
    Dim x2, y3, ResultPT As Double

    Private Sub ButExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButExit.Click
        Global.System.Windows.Forms.Application.Exit()
        Help.Visible = False
    End Sub

    Private Sub cb1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb1.Click
        If cb1.Checked = True Then

            gb4.Enabled = True
            txbx_tc.Enabled = True
            txbx_vsh.Enabled = True
            txbx_ccr.Enabled = True
            txbx_ct.Enabled = True
            txbx_c.Enabled = True
            txbx_av.Enabled = True
            txbx_ts.Enabled = True
            txbx_hte.Enabled = True
            txbx_pt.Enabled = True

            lbl_tc.Enabled = True
            lbl_vsh.Enabled = True
            lbl_ccr.Enabled = True
            lbl_ct.Enabled = True
            lbl_c.Enabled = True
            lbl_av.Enabled = True
            lbl_ts.Enabled = True
            lbl_hte.Enabled = True
            lbl_pt.Enabled = True
        Else
            gb4.Enabled = False
            txbx_tc.Enabled = False
            txbx_vsh.Enabled = False
            txbx_ccr.Enabled = False
            txbx_ct.Enabled = False
            txbx_c.Enabled = False
            txbx_av.Enabled = False
            txbx_ts.Enabled = False
            txbx_hte.Enabled = False
            txbx_pt.Enabled = False

            lbl_tc.Enabled = False
            lbl_vsh.Enabled = False
            lbl_ccr.Enabled = False
            lbl_ct.Enabled = False
            lbl_c.Enabled = False
            lbl_av.Enabled = False
            lbl_ts.Enabled = False
            lbl_hte.Enabled = False
            lbl_pt.Enabled = False

        End If
    End Sub

    Private Sub cb2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb2.Click
        If cb2.Checked = True Then
            gb5.Enabled = True
            txbx_mt.Enabled = True
            txbx_tc.Enabled = True
            txbx_vsh.Enabled = True
            txbx_hof.Enabled = True
            txbx_c.Enabled = True
            txbx_av.Enabled = True
            txbx_ts.Enabled = True
            txbx_p.Enabled = True
            ButRef.Enabled = True
            txbx_hte.Enabled = True

            lbl_mt.Enabled = True
            lbl_tc.Enabled = True
            lbl_vsh.Enabled = True
            lbl_hof.Enabled = True
            lbl_c.Enabled = True
            lbl_av.Enabled = True
            lbl_ts.Enabled = True
            lbl_p.Enabled = True
            lbl_hte.Enabled = True

        Else
            gb5.Enabled = False
            txbx_mt.Enabled = False
            txbx_tc.Enabled = False
            txbx_vsh.Enabled = False
            txbx_hof.Enabled = False
            txbx_c.Enabled = False
            txbx_av.Enabled = False
            txbx_ts.Enabled = False
            txbx_p.Enabled = False
            ButRef.Enabled = False
            txbx_hte.Enabled = False

            lbl_mt.Enabled = False
            lbl_tc.Enabled = False
            lbl_vsh.Enabled = False
            lbl_hof.Enabled = False
            lbl_c.Enabled = False
            lbl_av.Enabled = False
            lbl_ts.Enabled = False
            lbl_p.Enabled = False
            lbl_hte.Enabled = False


        End If
    End Sub

    Private Sub cb3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb3.Click
        If cb3.Checked = True Then
            gb6.Enabled = True
            txbx_mt.Enabled = True
            txbx_vsh.Enabled = True
            txbx_at.Enabled = True
            txbx_c.Enabled = True
            txbx_av.Enabled = True
            txbx_ts.Enabled = True
            txbx_p.Enabled = True
            ButRef.Enabled = True
            txbx_hte.Enabled = True
            txbx_pt.Enabled = True

            lbl_mt.Enabled = True
            lbl_vsh.Enabled = True
            lbl_at.Enabled = True
            lbl_c.Enabled = True
            lbl_av.Enabled = True
            lbl_ts.Enabled = True
            lbl_p.Enabled = True
            lbl_hte.Enabled = True
            lbl_pt.Enabled = True

        Else
            gb6.Enabled = False
            txbx_mt.Enabled = False
            txbx_vsh.Enabled = False
            txbx_at.Enabled = False
            txbx_c.Enabled = False
            txbx_av.Enabled = False
            txbx_ts.Enabled = False
            txbx_p.Enabled = False
            ButRef.Enabled = False
            txbx_hte.Enabled = False
            txbx_pt.Enabled = False

            lbl_mt.Enabled = False
            lbl_vsh.Enabled = False
            lbl_at.Enabled = False
            lbl_c.Enabled = False
            lbl_av.Enabled = False
            lbl_ts.Enabled = False
            lbl_p.Enabled = False
            lbl_hte.Enabled = False
            lbl_pt.Enabled = False

        End If
    End Sub

    Private Sub cb4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb4.Click
        If cb4.Checked = True Then
            gb7.Enabled = True
            txbx_mt.Enabled = True
            txbx_vsh.Enabled = True
            txbx_dfcl.Enabled = True
            txbx_c.Enabled = True
            txbx_av.Enabled = True
            txbx_ts.Enabled = True
            txbx_p.Enabled = True
            ButRef.Enabled = True
            txbx_hte.Enabled = True
            txbx_pt.Enabled = True

            lbl_mt.Enabled = True
            lbl_vsh.Enabled = True
            lbl_dfcl.Enabled = True
            lbl_c.Enabled = True
            lbl_av.Enabled = True
            lbl_ts.Enabled = True
            lbl_p.Enabled = True
            lbl_hte.Enabled = True
            lbl_pt.Enabled = True

        Else
            gb7.Enabled = False
            txbx_mt.Enabled = False
            txbx_vsh.Enabled = False
            txbx_dfcl.Enabled = False
            txbx_c.Enabled = False
            txbx_av.Enabled = False
            txbx_ts.Enabled = False
            txbx_p.Enabled = False
            ButRef.Enabled = False
            txbx_hte.Enabled = False
            txbx_pt.Enabled = False

            lbl_mt.Enabled = False
            lbl_vsh.Enabled = False
            lbl_dfcl.Enabled = False
            lbl_c.Enabled = False
            lbl_av.Enabled = False
            lbl_ts.Enabled = False
            lbl_p.Enabled = False
            lbl_hte.Enabled = False
            lbl_pt.Enabled = False


        End If
    End Sub

    Private Sub ButCal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButCal.Click
        mt = Val(txbx_mt.Text)
        tc = Val(txbx_tc.Text)
        vsh = Val(txbx_vsh.Text)
        ccr = Val(txbx_ccr.Text)
        ct = Val(txbx_ct.Text)
        at = Val(txbx_at.Text)
        hof = Val(txbx_hof.Text)
        dfcl = Val(txbx_dfcl.Text)
        c = Val(txbx_c.Text)
        av = Val(txbx_av.Text)
        ts = Val(txbx_ts.Text)
        p = Val(txbx_p.Text)
        hte = Val(txbx_hte.Text)
        pt = Val(txbx_pt.Text)
        Hnet = (c * av * hte) / ts

        tik = (Hnet * ccr) / (2 * System.Math.PI * tc)
        tik = (tik) ^ (1 / 2)
        tik = ct - tik
        TowTik = pt * ((vsh * (ct - tik) / Hnet) ^ (1 / 2))

        tin = ct - ((((Hnet ^ 2) * ccr) / (2 * System.Math.PI * tc * vsh * (pt ^ 2))) ^ (1 / 3))
        TowTin = pt * ((vsh * (ct - tin) / Hnet) ^ (1 / 2))

        If tik < 25 Or tin < 25 Then
            lbl1.Text = 25
            lbl2.Text = "Dusky"
        Else
            If TowTin < 0.75 Then
                lbl1.Text = System.Math.Round(tin, 0)
                lbl2.Text = "Thin Plate"
            End If
            If TowTik > 0.75 Then
                lbl1.Text = System.Math.Round(tik, 0)
                lbl2.Text = "Thick Plate"
            End If
            If TowTin > 0.75 And TowTik < 0.75 Then
                lbl1.Text = "Dusky"
                lbl2.Text = "Dusky"
            End If

        End If
       
        x = hof * Hnet
        y1 = 2 * System.Math.PI * tc * vsh
        y2 = (mt - p) ^ 2
        y = y1 * y2
        ResultST = System.Math.Round((x / y), 2)
        lbl3.Text = ResultST

        z = 1 / (at - p)
        x1 = 1 / (mt - p)
        w = (4.13 * vsh * pt) / Hnet
        ResultWoH = System.Math.Round(((z - x1) / w) * 2, 2)
        lbl5.Text = ResultWoH
        lbl4.Text = at

        x2 = ((p * 4.13 * vsh * pt * dfcl) / Hnet) + (p / (mt - p)) + 1
        y3 = ((4.13 * vsh * pt * dfcl) / Hnet) + (1 / (mt - p))
        ResultPT = System.Math.Round((x2 / y3), 0)
        lbl6.Text = ResultPT
        lbl7.Text = dfcl * 2

    End Sub

    Private Sub ButClr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButClr.Click
        txbx_mt.Text = ""
        txbx_tc.Text = ""
        txbx_vsh.Text = ""
        txbx_ccr.Text = ""
        txbx_ct.Text = ""
        txbx_at.Text = ""
        txbx_hof.Text = ""
        txbx_dfcl.Text = ""
        txbx_c.Text = ""
        txbx_av.Text = ""
        txbx_ts.Text = ""
        txbx_p.Text = ""
        txbx_hte.Text = ""
        txbx_pt.Text = ""
        lbl1.Text = "-------"
        lbl2.Text = "-------"
        lbl3.Text = "-------"
        lbl4.Text = "-------"
        lbl5.Text = "-------"
        lbl6.Text = "-------"
        lbl7.Text = "-------"

    End Sub

    Private Sub CalcuToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CalcuToolStripMenuItem.Click
        ButCal.PerformClick()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        ButExit.PerformClick()
    End Sub

    Private Sub WeldingConditionToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles WeldingConditionToolStripMenuItem.Click
        txbx_c.Text = ""
        txbx_av.Text = ""
        txbx_ts.Text = ""
        txbx_p.Text = ""
        txbx_hte.Text = ""
        txbx_pt.Text = ""
        lbl1.Text = "-------"
        lbl2.Text = "-------"
        lbl3.Text = "-------"
        lbl4.Text = "-------"
        lbl5.Text = "-------"
        lbl6.Text = "-------"
        lbl7.Text = "-------"
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click
        ButClr.PerformClick()
    End Sub

    Private Sub TherToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TherToolStripMenuItem.Click
        txbx_mt.Text = ""
        txbx_tc.Text = ""
        txbx_vsh.Text = ""
        txbx_ccr.Text = ""
        txbx_ct.Text = ""
        txbx_at.Text = ""
        txbx_hof.Text = ""
        txbx_dfcl.Text = ""
        lbl1.Text = "-------"
        lbl2.Text = "-------"
        lbl3.Text = "-------"
        lbl4.Text = "-------"
        lbl5.Text = "-------"
        lbl6.Text = "-------"
        lbl7.Text = "-------"
    End Sub

    Private Sub ThermalPropertiesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ThermalPropertiesToolStripMenuItem1.Click
        Help.Visible = True
        Me.Opacity = 50%
        Me.Enabled = False
    End Sub

    Private Sub ThermalPropertiesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ThermalPropertiesToolStripMenuItem.Click
        Convert.Visible = True
        Me.Opacity = 50%
        Me.Enabled = False
    End Sub

    Private Sub PersianToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PersianToolStripMenuItem.Click
        gb1.Text = "بخش محاسبات"
        gb2.Text = "خصوصیات حرارتی"
        gb3.Text = "شرایط جوشکاری"
        gb4.Text = "دمای پیش گرم"
        gb5.Text = "زمان انجماد"
        gb6.Text = "پهنای منطقه متاثر از حرارت"
        gb7.Text = "دما در فاصله مشخص از مرکز"
        cb1.Text = "تعیین دمای پیش گرم"
        cb2.Text = "تعیین زمان انجماد"
        cb3.Text = "تعیین پهنای منطقه متاثر از حرارت"
        cb4.Text = "تعیین بیشترین دما در موقعیت مشخص"
        lbl_mt.Text = "دمای ذوب (°C)"
        lbl_tc.Text = "ضریب هدایت حرارتی (J/mm.s.°C)"
        lbl_vsh.Text = "گرمای ویژه حجمی (J/mm³.°C)"
        lbl_ccr.Text = "نرخ سرد شدن بحرانی (°C/s )"
        lbl_ct.Text = "دمای بحرانی (°C )"
        lbl_at.Text = "دمای آستنیته (°C )"
        lbl_hof.Text = "گرمای نهان ذوب (J/mm³) "
        lbl_dfcl.Text = "فاصله از خط مرکز ( mm)"
        lbl_c.Text = "شدت جریان (A)"
        lbl_av.Text = "اختلاف پتانسیل ( V )"
        lbl_ts.Text = "سرعت حرکت ( mm/s )"
        lbl_p.Text = "دمای پیش گرم (°C) "
        lbl_hte.Text = "بازده انتقال حرارت"
        lbl_pt.Text = "ضخامت ورق ( mm )"
        ButClr.Text = "پاک کردن"
        ButCal.Text = "محاسبه"
        ButExit.Text = "خروج"
        FileToolStripMenuItem.Text = "فایل"
        CalcuToolStripMenuItem.Text = "محاسبه"
        ExitToolStripMenuItem.Text = "خروج"
        HelpToolStripMenuItem.Text = "ابزارها"
        ThermalPropertiesToolStripMenuItem.Text = "جدول تبدیل واحد"
        ClearAllToolStripMenuItem.Text = "پاک کردن"
        TherToolStripMenuItem.Text = "خصوصیات حرارتی"
        WeldingConditionToolStripMenuItem.Text = "شرایط جوشکاری"
        AllToolStripMenuItem.Text = "همه"
        HelpToolStripMenuItem1.Text = "راهنما"
        ThermalPropertiesToolStripMenuItem1.Text = "خصوصیات حرارتی"
        LanguageToolStripMenuItem.Text = "زبان"
        EnglishToolStripMenuItem.Text = "انگلیسی"
        PersianToolStripMenuItem.Text = "فارسی"
        Convert.But_swap.Text = "تعویض"
        Convert.But_close.Text = "بستن"
        Help.ButClose.Text = "بستن"
        PreRef.ButClose.Text = "بستن"
        lbl_pt_e.Visible = False
        lbl_pt_p.Visible = True
        lbl_ptR_e.Visible = False
        lbl_ptR_p.Visible = True
        lbl_st_p.Visible = True
        lbl_st_e.Visible = False
        lbl_woh_e.Visible = False
        lbl_woh_p.Visible = True
        lbl_ptasd_e.Visible = False
        lbl_ptasd_p.Visible = True
        EnglishToolStripMenuItem.Checked = False
        PersianToolStripMenuItem.Checked = True
        
    End Sub
    Private Sub EnglishToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnglishToolStripMenuItem.Click
        gb1.Text = "Calculate item"
        gb2.Text = "Thermal properties"
        gb3.Text = "Welding condition"
        gb4.Text = "Preheating temperatute"
        gb5.Text = "Solidification time"
        gb6.Text = "Width of HAZ"
        gb7.Text = "Peak temperature at specific distance of center line"
        cb1.Text = "Specify preheating temperatute"
        cb2.Text = "Specify solidification time"
        cb3.Text = "Specify width of HAZ"
        cb4.Text = "Specify peak temperature at specific locations of HAZ"
        lbl_mt.Text = "Melting temperature (°C)"
        lbl_tc.Text = "Thermal conductivity (J/mm.s.°C)"
        lbl_vsh.Text = "Volumetric specific heat (J/mm³.°C)"
        lbl_ccr.Text = "Critical cooling rate (°C/s )"
        lbl_ct.Text = "Critical temperature (°C)"
        lbl_at.Text = "Austenization temperature (°C)"
        lbl_hof.Text = "Heat of fusion (J/mm³) "
        lbl_dfcl.Text = "Distance from center line (mm)"
        lbl_c.Text = "Current (A)"
        lbl_av.Text = "Arc voltag (V)"
        lbl_ts.Text = "Travel speed (mm/s)"
        lbl_p.Text = "Preheat (°C) "
        lbl_hte.Text = "Heat-transfer efficiency"
        lbl_pt.Text = "Plate thickness (mm)"
        ButClr.Text = "Clear all"
        ButCal.Text = "Calculation"
        ButExit.Text = "Exit"
        FileToolStripMenuItem.Text = "File"
        CalcuToolStripMenuItem.Text = "Calculation"
        ExitToolStripMenuItem.Text = "Exit"
        HelpToolStripMenuItem.Text = "Tools"
        ThermalPropertiesToolStripMenuItem.Text = "Convert"
        ClearAllToolStripMenuItem.Text = "Clear all"
        TherToolStripMenuItem.Text = "Thermal properties"
        WeldingConditionToolStripMenuItem.Text = "Welding condition"
        AllToolStripMenuItem.Text = "All"
        HelpToolStripMenuItem1.Text = "Help"
        ThermalPropertiesToolStripMenuItem1.Text = "Thermal properties"
        LanguageToolStripMenuItem.Text = "Language"
        EnglishToolStripMenuItem.Text = "English"
        PersianToolStripMenuItem.Text = "Persian"
        Convert.But_swap.Text = "Swap"
        Convert.But_close.Text = "Close"
        Help.ButClose.Text = "Close"
        PreRef.ButClose.Text = "Close"
        lbl_pt_e.Visible = True
        lbl_pt_p.Visible = False
        lbl_ptR_e.Visible = True
        lbl_ptR_p.Visible = False
        lbl_st_p.Visible = False
        lbl_st_e.Visible = True
        lbl_woh_e.Visible = True
        lbl_woh_p.Visible = False
        lbl_ptasd_e.Visible = True
        lbl_ptasd_p.Visible = False
        EnglishToolStripMenuItem.Checked = True
        PersianToolStripMenuItem.Checked = False
    End Sub
    Private Sub ButRef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButRef.Click
        PreRef.Visible = True
    End Sub

    Private Sub ButDiag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButDiag.Click
        Dim distanceMultipler As Integer
        distanceMultipler = 3
        Dim a As Object
        Dim b As Microsoft.Office.Interop.Excel.Workbook
        Dim c As Microsoft.Office.Interop.Excel.Worksheet

        If lbl6.Text = "-------" Then
            If EnglishToolStripMenuItem.Checked = True Then
                MsgBox("Please first click calculation button")
            Else
                MsgBox("لطفا ابتدا دکمه محاسبه را فشار دهید ")
            End If

        Else
            a = Microsoft.VisualBasic.Interaction.CreateObject("Excel.Application")
            b = a.workbooks.open(My.Application.Info.DirectoryPath & "\Export Format.xlsx")
            c = b.Worksheets(1)

            x2 = ((p * 4.13 * vsh * pt * (0 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (0 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(3, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(3, 1) = 0 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (1 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (1 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(4, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(4, 1) = 1 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (2 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (2 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(5, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(5, 1) = 2 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (3 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (3 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(6, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(6, 1) = 3 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (4 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (4 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(7, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(7, 1) = 4 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (5 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (5 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(8, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(8, 1) = 5 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (6 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (6 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(9, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(9, 1) = 6 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (7 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (7 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(10, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(10, 1) = 7 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (8 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (8 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(11, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(11, 1) = 8 * distanceMultipler

            x2 = ((p * 4.13 * vsh * pt * (9 * distanceMultipler)) / Hnet) + (p / (mt - p)) + 1
            y3 = ((4.13 * vsh * pt * (9 * distanceMultipler)) / Hnet) + (1 / (mt - p))
            c.Cells(12, 2) = System.Math.Round((x2 / y3), 0)
            c.Cells(12, 1) = 9 * distanceMultipler


            a.visible = True
        End If

    End Sub
End Class