Imports System.Math
Public Class XMR

    Dim Vp, Vs, Vhv, Vlv, K, Ef, Fx, Ipri, Isec, Ihv, Ilv, Ahv, Alv, Bm, Ai, Agi, ki, d, Kw, Aw, J, D1, Ww, Hw, Hw_Ww_Ratio, a, b, c, e1, Ay, Dy, Hy, H, W, Shell_Const, S_a, S_depth, bare_diaL, bare_diaH As Double
    Dim Tm, f, Q, Thv, Tlv, Space_ReqL, Space_ReqH, spacersL, spacersH As Integer
    Dim Width_wiseWL, Height_wiseWL, Width_wiseCL, Height_wiseCL, clearanceL, axial_lenL, radial_depL, DiaInL, DiaOutL, MeanDiaL, MeanLenL, turnsperlyrL, HeightLV As Double
    Dim Width_wiseWH, Height_wiseWH, Width_wiseCH, Height_wiseCH, clearanceH, axial_lenH, radial_depH, DiaInH, DiaOutH, MeanDiaH, MeanLenH, turnsperlyrH, HeightHV As Double
    Dim r1, r2, Pc As Double

    Public Sub XMR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
       

        Q = Val(TextBox1.Text).ToString                        '//Power Rating of transformer in kVA
        Vp = Val(TextBox2.Text).ToString                       '//Primary side voltage in kV
        Vs = Val(TextBox3.Text).ToString                       '//Secondary side voltage in kV
        Tm = Val(TextBox4.Text).ToString                       '//Maximum temperature rise 
        Bm = Val(TextBox5.Text).ToString                       '//Maximum flux density in Wb/sq m
        J = (Val(TextBox6.Text) * 10 ^ 6).ToString             '//Curent Desity in A/sq mm


        TextBox1.Select()                                      '//Set focus on Textbox1 (Q) when form loads

        f = 50                                                 '//Frequency=50(Constant)
        ki = 0.9                                               '//Stacking factor(Assumed)


        '//Setting all combobox initial selection to first option in their respective lists
        ComboBox1.SelectedIndex = 0                             '//Core material(CRGO,HR...
        ComboBox2.SelectedIndex = 1                             '//Construction type(K)[Core type, Shell type]
        ComboBox3.SelectedIndex = 1                             '//Transformer type(Dist,Pwr...)

        
        Label52.Text = Val(TrackBar2.Value.ToString) / 100      '//Showing initial(Default) values of Lamination thickness

        If TabControl1.TabIndex = 1 Then                        '//On form load Tabpage2 is disabled 
            TabPage2.Enabled = False


        End If

    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

        'Only numbers & one decimal point can be entered in textbox 
        'There are a couple scenerios this code is looking for. One is checking for the Decimal period "."
        'and whether it exists in the textbox already. The other is seeing if the keypress was a Number and
        'Control based key.
        '
        If e.KeyChar = "." Then
            '
            'If a value higher than -1 is returned, it means there IS a existing decimal point’
            If TextBox1.Text.IndexOf(".") > -1 Then
                '
                'This says that I already dealt with the _KeyPress event so do not do anything else with this event.

                e.Handled = True

            End If
            '
            'Remember you want numbers to get through and the Control keys like “Backspace”
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            '
            'This says that I already dealt with the _KeyPress event so do not do anything else with this event.

            e.Handled = True

        End If

    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        '
        '//Only numbers & one decimal point validation
        If e.KeyChar = "." Then
            If TextBox2.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        '
        '//Only numbers & one decimal point validation

        If e.KeyChar = "." Then
            If TextBox3.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        '
        '//Only numbers & one decimal point validation

        If e.KeyChar = "." Then
            If TextBox4.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Val(TextBox1.Text) > 400 Then
            MessageBox.Show("Power rating of a 1-ph transformer beyond 400 kVA is not economical.Please enter the power rating of a 1- ph transformer below 400kVA .", "XMR", _
           MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox1.Text = ""
            TextBox1.Select()
            GoTo Proceed
        End If

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MessageBox.Show("Please enter the values in all the fields provided.", "XMR", _
           MessageBoxButtons.OK, MessageBoxIcon.Information)
            Button3.Enabled = False
            TextBox1.Select()
            GoTo Reset


        Else
            Button3.Enabled = True                                                  '//Clicking on Button1(Calculate) enables Button3(Proceed)
        End If


        If ComboBox2.SelectedItem = "Core Type" Then                                '//Setting value of K when Core type or Shell type is
            K = 0.8                                                                 'in selected in Contruction Type[K]
        Else
            K = 1.1
        End If

        Ef = K * Val(Math.Sqrt(TextBox1.Text))                                      '//Calculating EMF per turn and storing the value in 'Ef' in V
        Label10.Text = Ef.ToString("n4") & " V"                                            '//Ef is converted into string and displayed in label10


        Fx = Ef / (4.44 * f)                                                        '//Calculating Flux in Core and storing the value in Fx
        Label15.Text = Fx.ToString("n4") & " Wb"                                            '//Fx is converted into string and displayed in label 15

        Ipri = (Val(TextBox1.Text)) / (Val(TextBox2.Text))                          '//Primary Coil Current in A

        Isec = (Ipri * Val(TextBox2.Text)) / (Val(TextBox3.Text))                   '//Secondary Coil Current in A

        '//Coverting the values of current into strings
        Label24.Text = Ipri.ToString("n4") & " A"
        Label26.Text = Isec.ToString("n4") & " A"


        If Val(TextBox2.Text) > Val(TextBox3.Text) Then                             '//If Vp > Vs
            Label57.Text = Val(TextBox2.Text).ToString("n4")
            Label59.Text = Val(TextBox3.Text).ToString("n4")

            Vhv = Label57.Text.ToString                                             'Vhv=Vp
            Vlv = Label59.Text.ToString                                             'Vlv=Vs
        Else

            Label57.Text = Val(TextBox3.Text).ToString("n4")
            Label59.Text = Val(TextBox2.Text).ToString("n4")

            Vlv = Label57.Text.ToString
            Vhv = Label59.Text.ToString
        End If
        GoTo Proceed

Reset:
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""

        TextBox1.Select()
Proceed:
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.TabControl1.SelectedIndex = 1
        TabPage2.Enabled = True

    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar = "." Then
            If TextBox5.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub


    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If e.KeyChar = "." Then
            If TextBox6.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
       

        If TextBox5.Text = "" Then
            MessageBox.Show("Enter value of Flux Density", "XMR", _
                      MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox5.Select()
        Else
            Ai = (Val(Label15.Text) / Val(TextBox5.Text))              'Area in sq m
            Label32.Text = Ai.ToString("n4") & " sq m"

            Agi = Ai / ki
            Label34.Text = Agi.ToString("n4") & " sq m"
        End If
    End Sub

    '//Tabpage3(Core Design)
    Private Sub TabPage3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Enter

        If ComboBox2.SelectedIndex = 0 Then                                         'Shell Type selected
            Me.PictureBox2.Image = My.Resources.EI_type
            Me.PictureBox3.Image = My.Resources.LL_type_S
            Me.PictureBox4.Image = My.Resources.UT_type

            RadioButton1.Text = "EI-Type"
            RadioButton2.Text = "LL-Type"
            RadioButton3.Text = "UT-Type"



        Else                                                                        'Core Type selected
            Me.PictureBox2.Image = My.Resources.UI_type
            Me.PictureBox3.Image = My.Resources.LL_type
            Me.PictureBox4.Image = My.Resources.Mitred

            RadioButton1.Text = "UI-Type"
            RadioButton2.Text = "LL-Type"
            RadioButton3.Text = "Mitred-Type"


        End If

    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 And ComboBox4.SelectedIndex <> 0 Then
            MessageBox.Show("Shell type transformers >>> Rectangular cross-section only", "XMR", _
           MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox4.SelectedIndex = 0
        End If

        If ComboBox2.SelectedIndex = 1 And ComboBox4.SelectedIndex = 0 Then
            MessageBox.Show("Only shell type transformers >>> Rectangular cross-section", "XMR", _
          MessageBoxButtons.OK, MessageBoxIcon.Information)
            ComboBox4.SelectedIndex = 1
        End If
        If ComboBox4.SelectedIndex = 0 Then
            Me.PictureBox1.Image = My.Resources.Rect 
            Label38.Text = "0"
            Label37.Visible = False
            Label38.Visible = False

            GroupBox11.Visible = False
        Else
            Label37.Visible = True
            Label38.Visible = True

            GroupBox11.Visible = True

        End If

        If ComboBox4.SelectedIndex = 1 Then
            Me.PictureBox1.Image = My.Resources._1_step
            d = Math.Sqrt(Ai / 0.45)
            Label38.Text = d.ToString("n4") & " m"
            a = (Math.Sqrt(Val(Label38.Text) ^ 2 / 2)).ToString("n4")
            b = 0
            c = 0
            e1 = 0


        End If
        If ComboBox4.SelectedIndex = 2 Then
            Me.PictureBox1.Image = My.Resources._2_step
            d = Math.Sqrt(Ai / 0.56)
            Label38.Text = d.ToString("n4") & " m"
            a = 0.85 * Val(Label38.Text)
            b = 0.53 * Val(Label38.Text)
            c = 0
            e1 = 0

        End If
        If ComboBox4.SelectedIndex = 3 Then
            Me.PictureBox1.Image = My.Resources._3_step
            d = Math.Sqrt(Ai / 0.6)
            Label38.Text = d.ToString("n4") & " m"

            a = 0.9 * Val(Label38.Text)
            b = 0.7 * Val(Label38.Text)
            c = 0.42 * Val(Label38.Text)
            e1 = 0

        End If
        If ComboBox4.SelectedIndex = 4 Then
            Me.PictureBox1.Image = My.Resources._4_step
            d = Math.Sqrt(Ai / 0.62)
            Label38.Text = d.ToString("n4") & " m"
            a = 0.92 * Val(Label38.Text)
            b = 0.78 * Val(Label38.Text)
            c = 0.6 * Val(Label38.Text)
            e1 = 0.36 * Val(Label38.Text)

        End If

        Label80.Text = a.ToString & " m"
        Label81.Text = b.ToString & " m"
        Label82.Text = c.ToString & " m"
        Label83.Text = e1.ToString & " m"

    End Sub

    '//Selecting Lamination thickness and showing

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        Label52.Text = Val(TrackBar2.Value.ToString) / 100
    End Sub
    '//On clicking on a picture box the border style of the perticular picture box changes to 3D and no changes in other picture boxes border style. 
    'Respective Radio Button also gets selected

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        PictureBox2.BorderStyle = BorderStyle.Fixed3D
        PictureBox3.BorderStyle = BorderStyle.None
        PictureBox4.BorderStyle = BorderStyle.None
        RadioButton1.Checked = True
    End Sub
    '//On clicking on a picture box the border style of the perticular picture box changes to 3D and no changes in other picture boxes border style. 
    'Respective Radio Button also gets selected

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        PictureBox2.BorderStyle = BorderStyle.None
        PictureBox3.BorderStyle = BorderStyle.Fixed3D
        PictureBox4.BorderStyle = BorderStyle.None
        RadioButton2.Checked = True

    End Sub
    '//On clicking on a picture box the border style of the perticular picture box changes to 3D and no changes in other picture boxes border style. 
    'Respective Radio Button also gets selected

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        PictureBox2.BorderStyle = BorderStyle.None
        PictureBox3.BorderStyle = BorderStyle.None
        PictureBox4.BorderStyle = BorderStyle.Fixed3D
        RadioButton3.Checked = True

    End Sub

    '//TabPage4(Frame Design)

    Private Sub TabPage4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage4.Enter

        Label55.Text = Val(TextBox1.Text).ToString("n4") & " kVA"                        '//Power rating of Transformer

        If Val(Label55.Text) < 50 Then                                          '//Calculating Kw < 50 kVA
            Kw = 8 / (30 + Val(Label57.Text))
        End If

        If Val(Label55.Text) >= 50 And Val(Label55.Text) < 200 Then             '//Calculating 50 kVA < Kw < 200 kVA
            Kw = 10 / (30 + Val(Label57.Text))
        End If

        If Val(Label55.Text) >= 200 Then                                        '//Calculating Kw >= 200 kVA
            Kw = 12 / (30 + Val(Label57.Text))

        End If
        Label41.Text = Kw.ToString("n4")                                        '//Showing value of Kw according to the above conditions

       

        If ComboBox1.SelectedIndex = 0 Then
            Ay = Agi
        Else
            Ay = 1.2 * Agi
        End If

        '//Calculating Area of Window (Aw)
        Aw = (Val(TextBox1.Text) * 1000 / (2.22 * f * Val(TextBox5.Text) * (Val(TextBox6.Text) * 10 ^ 6) * Val(Label41.Text) * Val(Label32.Text))) * 10 ^ 6             'area in sq mm
        Label43.Text = Aw.ToString("n4") & " sq mm"

        '//Calculating parameters according to the selection done in Parameters I
        If ComboBox2.SelectedIndex = 0 Then                             '//If Shell type is selected 
            Me.PictureBox5.Image = My.Resources.ShellW
            Shell_Const = 2.5                                           '//Shell_Const=(S_depth/2*S_a)=2.5
            S_a = Math.Sqrt(Agi / (2.5 * 4))
            S_depth = 2.5 * 4 * S_a ^ 2

            Ay = Agi / 2
            Label79.Text = Ay.ToString("n4") & " Sq m"
            Label78.Text = S_depth.ToString("n4")

            Hy = Ay / S_depth
            Label72.Text = Hy.ToString("n4") & " m"


            Ww = Math.Sqrt(Aw / 3)
            Label68.Text = Ww.ToString("n4") & " mm"

            Hw = Aw / Ww
            Label70.Text = Hw.ToString("n4") & " mm"

            Label44.Text = 3

            H = Hw + 2 * Hy
            Label74.Text = H.ToString("n4")

            W = 2 * Ww + 4 * S_a
            Label76.Text = W.ToString("n4")

            GroupBox6.Visible = False

        Else                                                            '//If Core type is selected
            Me.PictureBox5.Image = My.Resources.CoreW

            Label62.Text = a.ToString("n4")

            d = Val(Label38.Text)                                       '//Diameter of Circumscribing Circle
            Label64.Text = d.ToString("n4")

            D1 = (Val(Label64.Text) * 1.7)                              '//Distance between center of adjacent limbs
            Label66.Text = D1.ToString("n4")

            Hw_Ww_Ratio = 4                                            '//Ratio of Height to width of the window
            Label44.Text = Hw_Ww_Ratio.ToString

            Ww = Math.Sqrt(Aw / Hw_Ww_Ratio)                            '//Width of the Window
            Label68.Text = Ww.ToString("n4")

            Hw = Aw / Ww                                                '//Height of the window
            Label70.Text = Hw.ToString("n4")

            Label79.Text = Ay.ToString("n4") & " Sq mm"

            Dy = a
            Label78.Text = Dy.ToString("n4")

            Hy = Ay / Dy
            Label72.Text = Hy.ToString("n4")

            H = Hw + 2 * Hy
            Label74.Text = H.ToString("n4")

            W = D1 + a
            Label76.Text = W.ToString("n4")

            GroupBox6.Visible = True

        End If

       

    End Sub

    '//TabPage2(Parameters II)
    Private Sub TabPage2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Enter

        Label135.Text = TextBox1.Text
        Label18.Text = ComboBox1.SelectedItem.ToString
        Label21.Text = ComboBox2.SelectedItem.ToString
        Label22.Text = ComboBox3.SelectedItem.ToString

    End Sub
   
   
    '//TabPage5(Winding Design)

    Private Sub TabPage5_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.Enter

        Tlv = (Val(Label59.Text) * 1000) / Val(Label10.Text)
        Label90.Text = Tlv.ToString & " Turns"

        Thv = (Val(Label57.Text) * 1000 * Tlv) / (Val(Label59.Text) * 1000)
        Label91.Text = Thv.ToString & " Turns"

        Ihv = (Val(TextBox1.Text)) / Val(Label57.Text)
        Label105.Text = Ihv.ToString("n4") & " A"

        Ilv = (Val(TextBox1.Text)) / Val(Label59.Text)
        Label99.Text = Ilv.ToString("n4") & " A"

        Alv = Ilv / Val(TextBox6.Text)
        Label95.Text = Alv.ToString("n4") & " Sq mm"

        Ahv = Ihv / Val(TextBox6.Text)
        Label106.Text = Ahv.ToString("n4") & " Sq mm"

        '//Type of windings used in HV side and in LV side

        '//For LV Side

        If Val(TextBox1.Text) <= 100 Or Val(Label59.Text) <= 0.44 Then
            Label101.Text = "Helical"
        End If

        If (Val(TextBox1.Text) > 100 And Val(TextBox1.Text) <= 1000) Or (Val(Label59.Text) > 0.44 And Val(Label59.Text) < 11) Then
            Label101.Text = "Helical or Multilayer Helix"
        End If

        If Val(TextBox1.Text) > 400 Or Val(Label59.Text) >= 11 Then
            Label101.Text = "Disc or Helical"
        End If


        '//For HV Side

        If Val(TextBox1.Text) <= 100 Or Val(Label57.Text) <= 11 Then
            Label107.Text = "Helical"
        End If

        If (Val(TextBox1.Text) > 100 And Val(TextBox1.Text) <= 400) Or (Val(Label57.Text) > 11 And Val(Label57.Text) <= 33) Then
            Label107.Text = "Multilayer Helix or Disc"
        End If

        If Val(TextBox1.Text) > 400 Or Val(Label57.Text) > 33 Then
            Label107.Text = "Disc or Multilayer Helix"
        End If

        '
        'Tappings on HV Side:-
        '
        Label114.Text = Thv + ((Thv * 5) / 100).ToString("n0") & " Turns"
        Label115.Text = Thv + ((Thv * 2.5) / 100).ToString("n0") & " Turns"
        Label116.Text = Thv - ((Thv * 5) / 100).ToString("n0") & " Turns"
        Label117.Text = Thv - ((Thv * 2.5) / 100).ToString("n0") & " Turns"
        '
        '------------------------------------------------------------------------------------------------------------------------------------------
        '// HV and LV Side windings
        '------------------------------------------------------------------------------------------------------------------------------------------
        'For LV side winding
        '
        If Val(Label95.Text) <= 5 Then
            bare_diaL = Math.Sqrt((4 * Val(Label95.Text)) / 3.1416)
            Label84.Text = bare_diaL.ToString("n4") & " mm"
            Label237.Text = 1

            Label239.Text = (Val(bare_diaL) + 1).ToString("n4") & " mm Diameter"
            Label235.Visible = False

            Width_wiseWL = (bare_diaL + 1)
            Height_wiseWL = (bare_diaL + 1)

            Label222.Text = Height_wiseWL.ToString("n4") & " mm"
            Label219.Text = Width_wiseWL.ToString("n4") & " mm"
        Else
            Dim i As Single
            Dim m As Single
            Dim Con_area As Single
            Dim A As Single
            Dim min_error As Single = 1000.0
            Dim Area_error As Single
            Dim final_thick As Single
            Dim final_width As Single
            Dim Cond_thick As Single() = {0.8, 0.9, 1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2, 2.2, 2.5, 2.8, 3, 3.2, 3.5, 3.8, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8, 9, 10}
            Dim Cond_width As Single() = {0.9, 1, 1.2, 1.4, 1.6, 1.8, 2.0, 2.2, 2.5, 2.8, 3, 3.2, 3.5, 3.8, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 19, 20, 22, 25}

            A = Val(Label95.Text)
            For i = 0 To Cond_thick.GetUpperBound(0)
                For m = 0 To Cond_width.GetUpperBound(0)
                    Con_area = Cond_thick(i) * Cond_width(m)
                    Area_error = Math.Abs(Con_area - A)

                    If Area_error < min_error Then

                        min_error = Area_error
                        final_thick = Cond_thick(i)
                        final_width = Cond_width(m)



                    End If
                Next m
            Next i

            If Val(Label95.Text) > 5 And Val(Label95.Text) <= 45 Then
                Label84.Text = final_thick.ToString & " mm x " & final_width.ToString & " mm"
                Label237.Text = 1
                Label239.Visible = True
                Label235.Visible = True
                Label239.Text = (Val(final_thick) + 0.5).ToString("n4") & " mm Thick"
                Label235.Text = (Val(final_width) + 0.5).ToString("n4") & " mm Width"

                Width_wiseWL = (Val(Label239.Text) + 0.5)
                Height_wiseWL = (Val(Label235.Text) + 0.5)

                Label222.Text = Height_wiseWL.ToString & " mm"
                Label219.Text = Width_wiseWL.ToString & " mm"
            Else
                A = (Val(Label95.Text) / 2)
                Label237.Text = 2

                For i = 0 To Cond_thick.GetUpperBound(0)
                    For m = 0 To Cond_width.GetUpperBound(0)
                        Con_area = Cond_thick(i) * Cond_width(m)
                        Area_error = Math.Abs(Con_area - A)

                        If Area_error < min_error Then

                            min_error = Area_error
                            final_thick = Cond_thick(i)
                            final_width = Cond_width(m)

                        End If
                    Next m
                Next i

                Label84.Text = final_thick.ToString & " mm x " & final_width.ToString & " mm"
                Label239.Visible = True
                Label235.Visible = True
                Label239.Text = (Val(final_thick) + 0.5).ToString("n4") & " mm Thick"
                Label235.Text = (Val(final_width) + 0.5).ToString("n4") & " mm Width"

                Height_wiseWL = (Val(Label239.Text) + 0.5) + (Val(Label239.Text) + 0.5)
                Width_wiseWL = (Val(Label235.Text) + 0.5)

                Label222.Text = Height_wiseWL.ToString & " mm"
                Label219.Text = Width_wiseWL.ToString & " mm"

            End If

        End If
        '
        '------------------------------------------------------------------------------------------------------------------------------------------
        'For HV side winding
        '
        If Val(Label106.Text) <= 5 Then
            bare_diaH = Math.Sqrt((4 * Val(Label106.Text)) / 3.1416)
            Label278.Text = bare_diaH.ToString("n4") & " mm"
            Label261.Text = 1
            Label258.Text = (Val(bare_diaH) + 1).ToString("n4") & " mm Diameter"
            Label257.Visible = False

            Width_wiseWH = (bare_diaH + 1)
            Height_wiseWH = (bare_diaH + 1)

            Label263.Text = Height_wiseWH.ToString("n4") & " mm"
            Label262.Text = Width_wiseWH.ToString("n4") & " mm"

        Else

            Dim i As Single
            Dim m As Single
            Dim Con_area As Single
            Dim A As Single
            Dim min_error As Single = 1000.0
            Dim Area_error As Single
            Dim final_thick As Single
            Dim final_width As Single
            Dim Cond_thick As Single() = {0.8, 0.9, 1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2, 2.2, 2.5, 2.8, 3, 3.2, 3.5, 3.8, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8, 9, 10}
            Dim Cond_width As Single() = {0.9, 1, 1.2, 1.4, 1.6, 1.8, 2.0, 2.2, 2.5, 2.8, 3, 3.2, 3.5, 3.8, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 19, 20, 22, 25}

            A = Val(Label106.Text)
            For i = 0 To Cond_thick.GetUpperBound(0)
                For m = 0 To Cond_width.GetUpperBound(0)
                    Con_area = Cond_thick(i) * Cond_width(m)
                    Area_error = Math.Abs(Con_area - A)

                    If Area_error < min_error Then

                        min_error = Area_error
                        final_thick = Cond_thick(i)
                        final_width = Cond_width(m)

                    End If
                Next m
            Next i
            If Val(Label106.Text) > 5 And Val(Label106.Text) <= 45 Then
                Label261.Text = 1
                Label278.Text = final_thick.ToString & " mm x " & final_width.ToString & " mm"

                Label258.Visible = True
                Label257.Visible = True
                Label258.Text = (Val(final_thick) + 0.5).ToString("n4") & " mm Thick"
                Label257.Text = (Val(final_width) + 0.5).ToString("n4") & " mm Width"

                Width_wiseWH = (Val(Label258.Text) + 0.5)
                Height_wiseWH = (Val(Label257.Text) + 0.5)

                Label263.Text = Height_wiseWH.ToString & " mm"
                Label262.Text = Width_wiseWH.ToString & " mm"
            Else
                A = Val(Label106.Text) / 2
                Label261.Text = 2

                For i = 0 To Cond_thick.GetUpperBound(0)
                    For m = 0 To Cond_width.GetUpperBound(0)
                        Con_area = Cond_thick(i) * Cond_width(m)
                        Area_error = Math.Abs(Con_area - A)

                        If Area_error < min_error Then

                            min_error = Area_error
                            final_thick = Cond_thick(i)
                            final_width = Cond_width(m)

                        End If
                    Next m
                Next i

                Label278.Text = final_thick.ToString & " mm x " & final_width.ToString & " mm"
                Label258.Visible = True
                Label257.Visible = True
                Label258.Text = (Val(final_thick) + 0.5).ToString("n4") & " mm Thick"
                Label257.Text = (Val(final_width) + 0.5).ToString("n4") & " mm Width"

                Height_wiseWH = (Val(Label258.Text) + 0.5) + (Val(Label258.Text) + 0.5)
                Width_wiseWH = (Val(Label257.Text) + 0.5)

                Label263.Text = Height_wiseWH.ToString & " mm"
                Label262.Text = Width_wiseWH.ToString & " mm"
            End If
        End If

        '-------------------------------------------------------------------------------------------------------------------------------------------
        'In LV Winding

        If Val(Label59.Text) <= 11 Then

            Label234.Text = "Helical Winding"

            GroupBox15.Enabled = True
            GroupBox15.Visible = True

            GroupBox16.Enabled = False
            GroupBox16.Visible = False


            turnsperlyrL = Tlv / Val(Label237.Text)

            If Val(Label237.Text) = 2 Then

                Width_wiseCL = Width_wiseWL + Width_wiseWL
                Height_wiseCL = Height_wiseWL
                Space_ReqL = (Tlv / 2) + 1
            Else
                Width_wiseCL = Width_wiseWL
                Height_wiseCL = Height_wiseWL
                Space_ReqL = Tlv + 1
            End If

            axial_lenL = Space_ReqL * Height_wiseCL
            clearanceL = (Val(Label70.Text) - axial_lenL) / 2
            radial_depL = (2 * Height_wiseWL) + 0.5               'insulation b/w layers = 0.5mm

            DiaInL = d + 2 * 1.5                                 'insulation b/w core and LV winding = 1.5mm
            DiaOutL = DiaInL + 2 * radial_depL

            MeanDiaL = (DiaInL + DiaOutL) / 2
            MeanLenL = 3.1416 * MeanDiaL

        Else

            Label178.Text = "Disc Winding"

            GroupBox15.Enabled = False
            GroupBox15.Visible = False

            GroupBox16.Enabled = True
            GroupBox16.Visible = True

            TrackBar4.Maximum = Tlv
            TrackBar4.Minimum = 1



        End If
        '
        'Display:-
        '
        Label142.Text = turnsperlyrL.ToString
        Label149.Text = Space_ReqL.ToString
        Label141.Text = axial_lenL.ToString("n2")
        Label139.Text = radial_depL.ToString("n2")
        Label138.Text = clearanceL.ToString("n2")
        Label137.Text = DiaInL.ToString("n2")
        Label123.Text = DiaOutL.ToString("n2")
        Label118.Text = MeanLenL.ToString("n2")

        '-------------------------------------------------------------------------------------------------------------------------------------------

        'In HV Winding

        If Val(Label57.Text) <= 11 Then

            Label284.Text = "Helical Winding"

            GroupBox17.Enabled = True
            GroupBox17.Visible = True

            GroupBox18.Enabled = False
            GroupBox18.Visible = False


            turnsperlyrH = Tlv / Val(Label261.Text)

            If Val(Label261.Text) = 2 Then

                Width_wiseCH = Width_wiseWH + Width_wiseWH
                Height_wiseCH = Height_wiseWH
                Space_ReqH = (Tlv / 2) + 1
            Else
                Width_wiseCH = Width_wiseWH
                Height_wiseCH = Height_wiseWH
                Space_ReqH = Tlv + 1
            End If
            axial_lenH = Space_ReqH * Height_wiseCH
            clearanceH = (Val(Label70.Text) - axial_lenH) / 2
            radial_depH = (2 * Height_wiseWH) + 0.5               'insulation b/w layers = 0.5mm

            DiaInH = d + 2 * 1.5                                 'insulation b/w core and LV winding = 1.5mm
            DiaOutH = DiaInH + 2 * radial_depH

            MeanDiaH = (DiaInH + DiaOutH) / 2
            MeanLenH = 3.1416 * MeanDiaH

        Else

            Label295.Text = "Disc Winding"

            GroupBox17.Enabled = False
            GroupBox17.Visible = False

            GroupBox18.Enabled = True
            GroupBox18.Visible = True

            TrackBar7.Maximum = Thv
            TrackBar7.Minimum = 1



        End If
        '
        'Display:-
        '
        Label275.Text = turnsperlyrH.ToString
        Label276.Text = Space_ReqH.ToString
        Label274.Text = axial_lenH.ToString("n2")
        Label273.Text = radial_depH.ToString("n2")
        Label272.Text = clearanceH.ToString("n2")
        Label268.Text = DiaInH.ToString("n2")
        Label267.Text = DiaOutH.ToString("n2")
        Label266.Text = MeanLenH.ToString("n2")

    End Sub

    Private Sub TrackBar4_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar4.Scroll

        Label175.Text = Val(TrackBar4.Value.ToString)
        Label174.Text = Math.Floor(Tlv / Val(Label175.Text))
        Label150.Text = Tlv - (Val(Label175.Text) * Val(Label174.Text))

        TrackBar5.Maximum = Val(Label174.Text)
        TrackBar5.Minimum = 1

    End Sub

    Private Sub TrackBar5_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar5.Scroll

        Label249.Text = TrackBar5.Value.ToString
        Label251.Text = Val(Label249.Text).ToString & " Height"
        Label253.Text = Math.Round(Val(Label174.Text) / Val(Label249.Text)).ToString & " Width"
        Label254.Text = Math.Abs(Val(Label174.Text) - (Val(Label251.Text) * Val(Label253.Text)))

        Label256.Text = (Val(Label251.Text) * Val(Label222.Text)).ToString & " High"
        Label255.Text = (Val(Label237.Text) * Val(Label219.Text) * Val(Label253.Text)).ToString & " Wide"

       
        spacersL = 10                      'Spacers b/w each coil =10mm

        HeightLV = (Val(Label175.Text) * Val(Label256.Text) + (Val(Label175.Text) - 1) * spacersL)

        DiaInL = d + 2 * 10                'Distance b/w core and LV Winding
        DiaOutL = DiaInL + 2 * Val(Label255.Text)
        MeanDiaL = (DiaInL + DiaOutL) / 2
        MeanLenL = 3.1416 * MeanDiaL

        Label137.Text = DiaInL.ToString("n2")
        Label123.Text = DiaOutL.ToString("n2")
        Label118.Text = MeanLenL.ToString("n2")

    End Sub

    Private Sub TrackBar7_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar7.Scroll

        Label299.Text = Val(TrackBar7.Value.ToString)
        Label298.Text = Math.Floor(Tlv / Val(Label299.Text))
        Label296.Text = Tlv - (Val(Label299.Text) * Val(Label298.Text))

        TrackBar6.Maximum = Val(Label298.Text)
        TrackBar6.Minimum = 1

    End Sub

    Private Sub TrackBar6_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar6.Scroll

        Label292.Text = TrackBar6.Value.ToString
        Label290.Text = Val(Label292.Text).ToString & " Height"
        Label289.Text = Math.Round(Val(Label298.Text) / Val(Label292.Text)).ToString & " Width"
        Label288.Text = Math.Abs(Val(Label298.Text) - (Val(Label290.Text) * Val(Label289.Text)))

        Label286.Text = (Val(Label290.Text) * Val(Label263.Text)).ToString & " High"
        Label285.Text = (Val(Label261.Text) * Val(Label262.Text) * Val(Label289.Text)).ToString & " Wide"


        spacersH = 10                      'Spacers b/w each coil =10mm

        HeightHV = (Val(Label299.Text) * Val(Label286.Text) + (Val(Label299.Text) - 1) * spacersH)

        DiaInH = d + 2 * 10                'Distance b/w core and LV Winding
        DiaOutH = DiaInH + 2 * Val(Label285.Text)
        MeanDiaH = (DiaInH + DiaOutH) / 2
        MeanLenH = 3.1416 * MeanDiaH

        Label268.Text = DiaInH.ToString("n2")
        Label267.Text = DiaOutH.ToString("n2")
        Label266.Text = MeanLenH.ToString("n2")

    End Sub

   
    '// Tabpage6 (Results page)
    '
    Private Sub TabPage6_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage6.Enter

        Label171.Text = TextBox1.Text
        Label172.Text = Label10.Text
        Label173.Text = Label15.Text
        Label176.Text = Label24.Text
        Label179.Text = Label26.Text
        Label180.Text = Label99.Text
        Label181.Text = Label105.Text

        Label182.Text = Label57.Text
        Label183.Text = Label59.Text
        Label184.Text = TextBox5.Text
        Label185.Text = TextBox6.Text

        r1 = (0.021 * MeanLenH * Thv) / Ahv
        r2 = (0.021 * MeanLenL * Tlv) / Alv

        Label305.Text = r1.ToString("n2") & " Ohms"
        Label306.Text = r2.ToString("n2") & " Ohms"

        Pc = (Ihv * Ihv) * r1 + (Ilv * Ilv) * r2
        Label310.Text = Pc.ToString("n2") & " Watts"

        Label186.Text = Label32.Text
        Label187.Text = Label34.Text
        Label188.Text = ComboBox4.SelectedText

        Label190.Text = Label38.Text


        Label193.Text = Label52.Text
        Label194.Text = Label41.Text
        Label195.Text = Label70.Text
        Label196.Text = Label68.Text
        Label197.Text = Label72.Text

        Label200.Text = Label78.Text
        Label201.Text = Label74.Text
        Label202.Text = Label76.Text
        Label203.Text = Label44.Text

        Label210.Text = Label90.Text
        Label211.Text = Label95.Text
        Label212.Text = Label237.Text
        Label213.Text = Label137.Text
        Label214.Text = Label123.Text
        Label215.Text = Label118.Text

        Label217.Text = Label91.Text
        Label218.Text = Label106.Text

        Label220.Text = Label261.Text

        Label221.Text = Label268.Text
        Label223.Text = Label267.Text
        Label224.Text = Label266.Text

        Label125.Text = ComboBox1.SelectedItem.ToString
        Label124.Text = ComboBox2.SelectedItem.ToString
        Label122.Text = ComboBox3.SelectedItem.ToString


        r1 = (0.021 * MeanLenH * Thv) / Ahv
        r2 = (0.021 * MeanLenL * Tlv) / Alv

        Label305.Text = r1.ToString("n2") & " Ohms"
        Label306.Text = r2.ToString("n2") & " Ohms"

        Pc = (Ihv * Ihv) * r1 + (Ilv * Ilv) * r2
        Label310.Text = Pc.ToString("n2") & " Watts"

    End Sub

   
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        If TextBox5.Text = "" Or TextBox6.Text = "" Then
            MessageBox.Show("Enter the values in the fields specified", "XMR", _
           MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox5.Text = ""
            TextBox5.Select()
            TextBox6.Text = ""
        Else
            Me.TabControl1.SelectedIndex = 2
            TabPage3.Enabled = True
        End If

       
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox6.Text = "" Then
            MessageBox.Show("Enter value of Current Density", "XMR", _
                      MessageBoxButtons.OK, MessageBoxIcon.Information)
            TextBox6.Select()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.TabControl1.SelectedIndex = 3
        TabPage4.Enabled = True
    End Sub

   
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.TabControl1.SelectedIndex = 4
        TabPage5.Enabled = True
    End Sub

   
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.TabControl1.SelectedIndex = 5
        TabPage6.Enabled = True
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Me.Close()
    End Sub

End Class


