Imports System.Math
Public Class Form1

    Dim Vp, Vs, Vhv, Vlv, K, Ef, Fx, Ipri, Isec, Ihv, Ilv, Ahv, Alv, Bm, Ai, Agi, ki, d, Kw, Aw, J, D1, Ww, Hw, Hw_Ww_Ratio, a, b, c, e1, Ay, Dy, Hy, H, W, Shell_Const, S_a, S_depth, bare_diaL, bare_diaH As Double
    Dim Tm, f, Q, Thv, Tlv, Space_ReqL, Space_ReqH, spacersL, spacersH As Integer
    Dim Width_wiseWL, Height_wiseWL, Width_wiseCL, Height_wiseCL, clearanceL, axial_lenL, radial_depL, DiaInL, DiaOutL, MeanDiaL, MeanLenL, turnsperlyrL, HeightLV As Double
    Dim Width_wiseWH, Height_wiseWH, Width_wiseCH, Height_wiseCH, clearanceH, axial_lenH, radial_depH, DiaInH, DiaOutH, MeanDiaH, MeanLenH, turnsperlyrH, HeightHV As Double
    Dim r1, r2, Pc As Double


    'Private ToolTip As New ToolTip()
    'Private Sub Q_Info_PictureBox_MouseHover(sender As Object, e As EventArgs) Handles Q_Info_PictureBox.MouseHover
    '    ToolTip.Show("Q should be less than 400 KVA", Q_Info_PictureBox)
    'End Sub

    'Private Sub Transformer_Type_Info_PictureBox_MouseHover(sender As Object, e As EventArgs) Handles Transformer_Type_Info_PictureBox.MouseHover
    '    ToolTip.Show("Distribution Transformer: (Q upto 200 kVA) || Power Transformer: (Q greater than 200 kVA)", Transformer_Type_Info_PictureBox)
    'End Sub

    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Q_Input_TextBox.Select()                               '//Set focus on Q_Input_TextBox (Q) when form loads
        Q = Val(Q_Input_TextBox.Text).ToString                 '//Power Rating of transformer in kVA
        Vp = Val(Vp_Input_TextBox.Text).ToString               '//Primary side voltage in kV
        Vs = Val(Vs_Input_TextBox.Text).ToString               '//Secondary side voltage in kV
        Tm = Val(Deg_C_Input_TextBox.Text).ToString            '//Maximum temperature rise 
        Bm = Val(Bm_Input_TextBox.Text).ToString               '//Maximum flux density in Wb/sq m
        J = (Val(Current_Density_Input_TextBox.Text) * 10 ^ 6).ToString             '//Curent Desity in A/sq mm

        f = 50                                                 '//Frequency=50(Constant)
        ki = 0.9                                               '//Stacking factor(Assumed)

        '//Setting all combobox initial selection to first option in their respective lists
        Core_Material_Select_ComboBox.SelectedIndex = 0         '//Core material(CRGO,HR...
        K_Select_ComboBox.SelectedIndex = 0                     '//Construction type(K)[Core type, Shell type]
        Transformer_Type_Select_ComboBox.SelectedIndex = 0      '//Transformer type(Dist,Pwr...)

        Lamination_Thick_Trackbar_Label.Text = Val(Lamination_Thickness_TrackBar.Value.ToString) / 100      '//Showing initial(Default) values of Lamination thickness

        If TabControl1.TabIndex = 1 Then
            TabPage2.Enabled = False
            TabPage3.Enabled = False
            TabPage4.Enabled = False
            TabPage5.Enabled = False
        End If
    End Sub

    Private Sub Current_Density_Info_PictureBox_Click(sender As Object, e As EventArgs) Handles Current_Density_Info_PictureBox.Click
        Current_Density_Info_GroupBox.Visible = True
    End Sub

    Private Sub Bm_Info_PictureBox_Click(sender As Object, e As EventArgs) Handles Bm_Info_PictureBox.Click
        Bm_Info_GroupBox.Visible = True
    End Sub

    Private Sub Transformer_Type_Info_PictureBox_Click(sender As Object, e As EventArgs) Handles Transformer_Type_Info_PictureBox.Click
        Transformer_Type_Info_GroupBox.Visible = True
    End Sub

    Private Sub Q_Info_PictureBox_Click(sender As Object, e As EventArgs) Handles Q_Info_PictureBox.Click
        Q_Info_GroupBox.Visible = True
    End Sub

    Private Sub Q_Input_TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Q_Input_TextBox.KeyPress

        'Only numbers & one decimal point can be entered in textbox 
        'There are a couple scenerios this code is looking for. One is checking for the Decimal period "."
        'and whether it exists in the textbox already. The other is seeing if the keypress was a Number and
        'Control based key.
        '
        If e.KeyChar = "." Then
            '
            'If a value higher than -1 is returned, it means there IS a existing decimal point’
            If Q_Input_TextBox.Text.IndexOf(".") > -1 Then
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

    Private Sub Vp_Input_TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Vp_Input_TextBox.KeyPress
        '
        '//Only numbers & one decimal point validation
        If e.KeyChar = "." Then
            If Vp_Input_TextBox.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub Vs_Input_TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Vs_Input_TextBox.KeyPress
        '
        '//Only numbers & one decimal point validation

        If e.KeyChar = "." Then
            If Vs_Input_TextBox.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub Deg_C_Input_TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Deg_C_Input_TextBox.KeyPress
        '
        '//Only numbers & one decimal point validation

        If e.KeyChar = "." Then
            If Deg_C_Input_TextBox.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub Bm_Input_TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Bm_Input_TextBox.KeyPress
        If e.KeyChar = "." Then
            If Bm_Input_TextBox.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub


    Private Sub Current_Density_Input_TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Current_Density_Input_TextBox.KeyPress
        If e.KeyChar = "." Then
            If Current_Density_Input_TextBox.Text.IndexOf(".") > -1 Then
                e.Handled = True
            End If
        ElseIf Char.IsNumber(e.KeyChar) = False AndAlso Char.IsControl(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub Input_Parameter_Tab_Reset_Button_Click(sender As Object, e As EventArgs) Handles Input_Parameter_Tab_Reset_Button.Click
        Q_Input_TextBox.Enabled = True
        Vp_Input_TextBox.Enabled = True
        Vs_Input_TextBox.Enabled = True
        Deg_C_Input_TextBox.Enabled = True
        Bm_Input_TextBox.Enabled = True
        Current_Density_Input_TextBox.Enabled = True
        Core_Material_Select_ComboBox.Enabled = True
        K_Select_ComboBox.Enabled = True
        Transformer_Type_Select_ComboBox.Enabled = True
        Bm_Input_TextBox.Enabled = True
        Current_Density_Input_TextBox.Enabled = True

        Q_Input_TextBox.Text = ""
        Vp_Input_TextBox.Text = ""
        Vs_Input_TextBox.Text = ""
        Deg_C_Input_TextBox.Text = ""
        Bm_Input_TextBox.Text = ""
        Current_Density_Input_TextBox.Text = ""
        Core_Material_Select_ComboBox.SelectedIndex = 0
        K_Select_ComboBox.SelectedIndex = 0
        Transformer_Type_Select_ComboBox.SelectedIndex = 0
        Bm_Input_TextBox.Text = ""
        Current_Density_Input_TextBox.Text = ""

        Input_Parameter_Tab_Calculate_Button.Enabled = False
        Input_Parameter_Tab_Next_Button.Enabled = False
        Input_Parameter_Tab_Validate_Button.Enabled = True

        Calc_Result_GroupBox.Visible = False
    End Sub


    Private Sub Input_Parameter_Tab_Validate_Button_Click(sender As Object, e As EventArgs) Handles Input_Parameter_Tab_Validate_Button.Click

        If Q_Input_TextBox.Text = "" Or
            Vp_Input_TextBox.Text = "" Or
            Vs_Input_TextBox.Text = "" Or
            Deg_C_Input_TextBox.Text = "" Or
            Bm_Input_TextBox.Text = "" Or
            Current_Density_Input_TextBox.Text = "" Or
            Core_Material_Select_ComboBox.SelectedIndex = 0 Or
            K_Select_ComboBox.SelectedIndex = 0 Or
            Transformer_Type_Select_ComboBox.SelectedIndex = 0 Or
            Bm_Input_TextBox.Text = "" Or
            Current_Density_Input_TextBox.Text = "" Then

            MessageBox.Show("Please enter the values in all the fields provided.",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

        ElseIf Val(Q_Input_TextBox.Text) > 400 Then
            MessageBox.Show("Power rating of a 1-ph transformer beyond 400 kVA is not economical. Please enter the power rating of a 1- ph transformer below 400kVA .",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Q_Input_TextBox.Text = ""
            Q_Input_TextBox.Select()

        ElseIf Val(Q_Input_TextBox.Text) <= 0 Or
            Val(Vp_Input_TextBox.Text) <= 0 Or
            Val(Vs_Input_TextBox.Text) <= 0 Or
            Val(Bm_Input_TextBox.Text) <= 0 Or
            Val(Current_Density_Input_TextBox.Text) <= 0 Then
            MessageBox.Show("Value should be greater than 0",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
        ElseIf Val(Current_Density_Input_TextBox.Text) < 1.1 Or Val(Current_Density_Input_TextBox.Text) > 6.2 Then
            MessageBox.Show("Value of Current Density should be between 1.1 to 6.2, depends upon type of cooling method used",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Current_Density_Input_TextBox.Text = ""
            Current_Density_Input_TextBox.Select()

        Else
            MessageBox.Show("All fields are validated. Go ahead for calculation",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

            Q_Input_TextBox.Enabled = False
            Vp_Input_TextBox.Enabled = False
            Vs_Input_TextBox.Enabled = False
            Deg_C_Input_TextBox.Enabled = False
            Bm_Input_TextBox.Enabled = False
            Current_Density_Input_TextBox.Enabled = False
            Core_Material_Select_ComboBox.Enabled = False
            K_Select_ComboBox.Enabled = False
            Transformer_Type_Select_ComboBox.Enabled = False
            Bm_Input_TextBox.Enabled = False
            Current_Density_Input_TextBox.Enabled = False
            Input_Parameter_Tab_Calculate_Button.Enabled = True
            Input_Parameter_Tab_Validate_Button.Enabled = False
        End If

    End Sub


    Private Sub Input_Parameter_Tab_Calculate_Button_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Input_Parameter_Tab_Calculate_Button.Click

        If K_Select_ComboBox.SelectedItem = "Core Type" Then                        '//Setting value of K when Core type or Shell type is
            K = 0.8                                                                 'in selected in Contruction Type[K]
        ElseIf K_Select_ComboBox.SelectedItem = "Shell Type" Then
            K = 1.1
        End If

        Ef = K * Val(Math.Sqrt(Q_Input_TextBox.Text))                               '//Calculating EMF per turn and storing the value in 'Ef' in V
        Voltage_Per_Turn_Calculated_Label.Text = Ef.ToString("n4")                  '//Ef is converted into string and displayed in label10

        Fx = Ef / (4.44 * f)                                                        '//Calculating Flux in Core and storing the value in Fx
        Flux_In_Core_Calculated_Label.Text = Fx.ToString("n4")                      '//Fx is converted into string and displayed in label 15

        Ai = Fx / Val(Bm_Input_TextBox.Text)                                        'Area in sq m
        Net_Core_Area_Calculated_Label.Text = Ai.ToString("n4")

        Agi = Ai / ki
        Gross_Core_Area_Calculated_Label.Text = Agi.ToString("n4")

        Ipri = (Val(Q_Input_TextBox.Text)) / (Val(Vp_Input_TextBox.Text))           '//Primary Coil Current in A
        Isec = (Ipri * Val(Vp_Input_TextBox.Text)) / (Val(Vs_Input_TextBox.Text))   '//Secondary Coil Current in A

        '//Coverting the values of current into strings
        Ip_Calculated_Label.Text = Ipri.ToString("n4")
        Is_Calculated_Label.Text = Isec.ToString("n4")

        If Val(Vp_Input_TextBox.Text) > Val(Vs_Input_TextBox.Text) Then             '//If Vp > Vs
            Vhv = Val(Vp_Input_TextBox.Text).ToString("n4")
            Vlv = Val(Vs_Input_TextBox.Text).ToString("n4")
        Else
            Vhv = Val(Vs_Input_TextBox.Text).ToString("n4")
            Vlv = Val(Vp_Input_TextBox.Text).ToString("n4")
        End If

        '//Coverting the values of voltages into strings
        Vhv_Calculated_Label.Text = Vhv.ToString("n4")
        Vlv_Calculated_Label.Text = Vlv.ToString("n4")

        Calc_Result_GroupBox.Visible = True
        Input_Parameter_Tab_Next_Button.Enabled = True       '//Clicking on Input_Parameter_Tab_Calculate_Button(Calculate) enables Input_Parameter_Tab_Next_Button(Next)
        Input_Parameter_Tab_Calculate_Button.Enabled = False
    End Sub

    Private Sub Input_Parameter_Tab_Next_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Input_Parameter_Tab_Next_Button.Click
        Me.TabControl1.SelectedIndex = 1
        TabPage2.Enabled = True
    End Sub





    '//Tabpage2(Core Design)
    Private Sub TabPage2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Enter

        If K_Select_ComboBox.SelectedItem = "Shell Type" Then
            Cross_Sec_Transformer_ComboBox.SelectedItem = "Rectangular Core"
            Cross_Sec_Transformer_ComboBox.Enabled = False
            Cross_Sec_Transformer_PictureBox.Visible = True
        Else
            Cross_Sec_Transformer_ComboBox.SelectedIndex = 0
            Cross_Sec_Transformer_ComboBox.Enabled = Enabled
        End If

        If K_Select_ComboBox.SelectedIndex = 1 Then                                 'Shell Type selected
            Me.PictureBox2.Image = My.Resources.EI_type
            Me.PictureBox3.Image = My.Resources.LL_type_S
            Me.PictureBox4.Image = My.Resources.UT_type

            Lamination_Type_RadioButton1.Text = "EI-Type"
            Lamination_Type_RadioButton2.Text = "LL-Type"
            Lamination_Type_RadioButton3.Text = "UT-Type"

        ElseIf K_Select_ComboBox.SelectedIndex = 2 Then                             'Core Type selected
            Me.PictureBox2.Image = My.Resources.UI_type
            Me.PictureBox3.Image = My.Resources.LL_type
            Me.PictureBox4.Image = My.Resources.Mitred

            Lamination_Type_RadioButton1.Text = "UI-Type"
            Lamination_Type_RadioButton2.Text = "LL-Type"
            Lamination_Type_RadioButton3.Text = "Mitred-Type"
        End If

        If Cross_Sec_Transformer_ComboBox.Enabled = True Then
            Cross_Sec_Transformer_PictureBox.Visible = False
        End If
    End Sub

    Private Sub Cross_Sec_Transformer_ComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cross_Sec_Transformer_ComboBox.SelectedIndexChanged

        Cross_Sec_Transformer_PictureBox.Visible = True
        Cross_Sec_Dimension_GroupBox.Visible = True

        If K_Select_ComboBox.SelectedItem = "Shell Type" And Cross_Sec_Transformer_ComboBox.SelectedItem <> "Rectangular Core" Then
            MessageBox.Show("For Shell Type Transformers: Rectangular Cross-Section Only",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Cross_Sec_Transformer_ComboBox.SelectedItem = "Rectangular Core"
        End If

        If K_Select_ComboBox.SelectedItem = "Core Type" And Cross_Sec_Transformer_ComboBox.SelectedItem = "Rectangular Core" Then
            MessageBox.Show("For Shell Type Transformers: Rectangular Cross-Section Only",
                            "Electrical Transformer Design",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
            Cross_Sec_Transformer_ComboBox.SelectedIndex = 0
        End If

        If Cross_Sec_Transformer_ComboBox.SelectedIndex = 1 Then
            Me.Cross_Sec_Transformer_PictureBox.Image = My.Resources.Rect
            Circumcircle_Diameter_Label.Visible = False
            Circumcircle_Diameter_Calculated_Label.Visible = False
            Cross_Sec_Dimension_GroupBox.Visible = False
        Else
            Circumcircle_Diameter_Label.Visible = True
            Circumcircle_Diameter_Calculated_Label.Visible = True
            Cross_Sec_Dimension_GroupBox.Visible = True
        End If

        If Cross_Sec_Transformer_ComboBox.SelectedIndex = 2 Then
            Me.Cross_Sec_Transformer_PictureBox.Image = My.Resources._1_step
            d = Math.Sqrt(Ai / 0.45)
            a = Math.Sqrt(d ^ 2 / 2)
            b = 0
            c = 0
            e1 = 0

        ElseIf Cross_Sec_Transformer_ComboBox.SelectedIndex = 3 Then
            Me.Cross_Sec_Transformer_PictureBox.Image = My.Resources._2_step
            d = Math.Sqrt(Ai / 0.56)
            a = 0.85 * d
            b = 0.53 * d
            c = 0
            e1 = 0

        ElseIf Cross_Sec_Transformer_ComboBox.SelectedIndex = 4 Then
            Me.Cross_Sec_Transformer_PictureBox.Image = My.Resources._3_step
            d = Math.Sqrt(Ai / 0.6)
            a = 0.9 * d
            b = 0.7 * d
            c = 0.42 * d
            e1 = 0

        ElseIf Cross_Sec_Transformer_ComboBox.SelectedIndex = 5 Then
            Me.Cross_Sec_Transformer_PictureBox.Image = My.Resources._4_step
            d = Math.Sqrt(Ai / 0.62)
            a = 0.92 * d
            b = 0.78 * d
            c = 0.6 * d
            e1 = 0.36 * d

        End If

        Label_a.Text = a.ToString("n4") & " m"
        Label_b.Text = b.ToString("n4") & " m"
        Label_c.Text = c.ToString("n4") & " m"
        Label_e1.Text = e1.ToString("n4") & " m"
        Circumcircle_Diameter_Calculated_Label.Text = d.ToString("n4") & " m"

    End Sub

    '//Selecting Lamination thickness and showing

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lamination_Thickness_TrackBar.Scroll
        Lamination_Thick_Trackbar_Label.Text = Val(Lamination_Thickness_TrackBar.Value.ToString) / 100
    End Sub
    '//On clicking on a picture box the border style of the perticular picture box changes to 3D and no changes in other picture boxes border style. 
    'Respective Radio Button also gets selected

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        PictureBox2.BorderStyle = BorderStyle.Fixed3D
        PictureBox3.BorderStyle = BorderStyle.None
        PictureBox4.BorderStyle = BorderStyle.None
        Lamination_Type_RadioButton1.Checked = True
    End Sub
    '//On clicking on a picture box the border style of the perticular picture box changes to 3D and no changes in other picture boxes border style. 
    'Respective Radio Button also gets selected

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        PictureBox2.BorderStyle = BorderStyle.None
        PictureBox3.BorderStyle = BorderStyle.Fixed3D
        PictureBox4.BorderStyle = BorderStyle.None
        Lamination_Type_RadioButton2.Checked = True

    End Sub
    '//On clicking on a picture box the border style of the perticular picture box changes to 3D and no changes in other picture boxes border style. 
    'Respective Radio Button also gets selected

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        PictureBox2.BorderStyle = BorderStyle.None
        PictureBox3.BorderStyle = BorderStyle.None
        PictureBox4.BorderStyle = BorderStyle.Fixed3D
        Lamination_Type_RadioButton3.Checked = True

    End Sub







    '//TabPage3(Frame Design)

    Private Sub TabPage3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Enter

        Q_Label_Copy_1.Text = Val(Q_Input_TextBox.Text).ToString("n4") & " kVA"        '//Power rating of Transformer

        If Val(Q_Input_TextBox.Text) < 50 Then                                  '//Calculating Kw < 50 kVA
            Kw = 8 / (30 + Val(Vhv_Calculated_Label.Text))
        End If

        If 50 <= Val(Q_Input_TextBox.Text) < 200 Then                           '//Calculating 50 kVA < Kw < 200 kVA
            Kw = 10 / (30 + Val(Vhv_Calculated_Label.Text))
        End If

        If Val(Q_Input_TextBox.Text) >= 200 Then                                '//Calculating Kw >= 200 kVA
            Kw = 12 / (30 + Val(Vhv_Calculated_Label.Text))

        End If
        Label41.Text = Kw.ToString("n4")                                        '//Showing value of Kw according to the above conditions



        If Core_Material_Select_ComboBox.SelectedIndex = 1 Then
            Ay = Agi
        ElseIf Core_Material_Select_ComboBox.SelectedIndex = 2 Then
            Ay = 1.2 * Agi
        End If

        '//Calculating Area of Window (Aw)
        Aw = (Val(Q_Input_TextBox.Text) * 1000 / (2.22 * f * Val(Bm_Input_TextBox.Text) * (Val(Current_Density_Input_TextBox.Text) * 10 ^ 6) * Val(Label41.Text) * Val(Net_Core_Area_Calculated_Label.Text))) * 10 ^ 6             'area in sq mm
        Label43.Text = Aw.ToString("n4") & " sq mm"

        '//Calculating parameters according to the selection done in Parameters I
        If K_Select_ComboBox.SelectedIndex = 1 Then                     '//If Shell type is selected 
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

        ElseIf K_Select_ComboBox.SelectedIndex = 2 Then                 '//If Core type is selected

            Me.PictureBox5.Image = My.Resources.CoreW

            Label62.Text = a.ToString("n4")

            d = Val(Circumcircle_Diameter_Calculated_Label.Text)        '//Diameter of Circumscribing Circle
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










    '//TabPage4(Winding Design)

    Private Sub TabPage5_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage4.Enter
        If TabPage4.Enabled = True Then
            Tlv = (Val(Vlv_Calculated_Label.Text) * 1000) / Val(Voltage_Per_Turn_Calculated_Label.Text)
            Label90.Text = Tlv.ToString & " Turns"
            Thv = (Val(Vhv_Calculated_Label.Text) * 1000 * Tlv) / (Val(Vlv_Calculated_Label.Text) * 1000)
            Label91.Text = Thv.ToString & " Turns"
            Ihv = (Val(Q_Input_TextBox.Text)) / Val(Vhv_Calculated_Label.Text)
            Label105.Text = Ihv.ToString("n4") & " A"
            Ilv = (Val(Q_Input_TextBox.Text)) / Val(Vlv_Calculated_Label.Text)
            Label99.Text = Ilv.ToString("n4") & " A"
            Alv = Ilv / Val(Current_Density_Input_TextBox.Text)
            Label95.Text = Alv.ToString("n4") & " Sq mm"
            Ahv = Ihv / Val(Current_Density_Input_TextBox.Text)
            Label106.Text = Ahv.ToString("n4") & " Sq mm"
        End If
        '//Type of windings used in HV side and in LV side

        '//For LV Side

        If Val(Q_Input_TextBox.Text) <= 100 Or Val(Vlv_Calculated_Label.Text) <= 0.44 Then
            Label101.Text = "Helical"
        End If

        If (Val(Q_Input_TextBox.Text) > 100 And Val(Q_Input_TextBox.Text) <= 1000) Or (Val(Vlv_Calculated_Label.Text) > 0.44 And Val(Vlv_Calculated_Label.Text) < 11) Then
            Label101.Text = "Helical or Multilayer Helix"
        End If

        If Val(Q_Input_TextBox.Text) > 400 Or Val(Vlv_Calculated_Label.Text) >= 11 Then
            Label101.Text = "Disc or Helical"
        End If


        '//For HV Side

        If Val(Q_Input_TextBox.Text) <= 100 Or Val(Vhv_Calculated_Label.Text) <= 11 Then
            Label107.Text = "Helical"
        End If

        If (Val(Q_Input_TextBox.Text) > 100 And Val(Q_Input_TextBox.Text) <= 400) Or (Val(Vhv_Calculated_Label.Text) > 11 And Val(Vhv_Calculated_Label.Text) <= 33) Then
            Label107.Text = "Multilayer Helix or Disc"
        End If

        If Val(Q_Input_TextBox.Text) > 400 Or Val(Vhv_Calculated_Label.Text) > 33 Then
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

        If Val(Vlv_Calculated_Label.Text) <= 11 Then

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

        If Val(Vhv_Calculated_Label.Text) <= 11 Then

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
    Private Sub TabPage6_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.Enter

        Label171.Text = Q_Input_TextBox.Text
        Label172.Text = Voltage_Per_Turn_Calculated_Label.Text
        Label173.Text = Flux_In_Core_Calculated_Label.Text
        Label176.Text = Ip_Calculated_Label.Text
        Label179.Text = Is_Calculated_Label.Text
        Label180.Text = Label99.Text
        Label181.Text = Label105.Text

        Label182.Text = Vhv_Calculated_Label.Text
        Label183.Text = Vlv_Calculated_Label.Text
        Label184.Text = Bm_Input_TextBox.Text
        Label185.Text = Current_Density_Input_TextBox.Text

        r1 = (0.021 * MeanLenH * Thv) / Ahv
        r2 = (0.021 * MeanLenL * Tlv) / Alv

        Label305.Text = r1.ToString("n2") & " Ohms"
        Label306.Text = r2.ToString("n2") & " Ohms"

        Pc = (Ihv * Ihv) * r1 + (Ilv * Ilv) * r2
        Label310.Text = Pc.ToString("n2") & " Watts"

        Label186.Text = Net_Core_Area_Calculated_Label.Text
        Label187.Text = Gross_Core_Area_Calculated_Label.Text
        Label188.Text = Cross_Sec_Transformer_ComboBox.SelectedText

        Label190.Text = Circumcircle_Diameter_Calculated_Label.Text


        Label193.Text = Lamination_Thick_Trackbar_Label.Text
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

        Label125.Text = Core_Material_Select_ComboBox.SelectedItem.ToString
        Label124.Text = K_Select_ComboBox.SelectedItem.ToString
        Label122.Text = Transformer_Type_Select_ComboBox.SelectedItem.ToString


        r1 = (0.021 * MeanLenH * Thv) / Ahv
        r2 = (0.021 * MeanLenL * Tlv) / Alv

        Label305.Text = r1.ToString("n2") & " Ohms"
        Label306.Text = r2.ToString("n2") & " Ohms"

        Pc = (Ihv * Ihv) * r1 + (Ilv * Ilv) * r2
        Label310.Text = Pc.ToString("n2") & " Watts"

    End Sub


    Private Sub Core_Design_Tab_Next_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Core_Design_Tab_Next_Button.Click
        Me.TabControl1.SelectedIndex = 2
        TabPage3.Enabled = True
    End Sub


    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.TabControl1.SelectedIndex = 4
        TabPage4.Enabled = True
    End Sub


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.TabControl1.SelectedIndex = 5
        TabPage5.Enabled = True
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Me.Close()
    End Sub

End Class


