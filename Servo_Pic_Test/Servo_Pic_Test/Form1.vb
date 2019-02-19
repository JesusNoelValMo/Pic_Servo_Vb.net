Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Nummod As Integer
        Nummod = NmcInit("COM6", 19200)
        If Nummod = 0 Then
            MsgBox("No hay modulos conectados")
        Else
            MsgBox("OK")
            addr = 1
            setGains(1, kp01, kd01, ki01, EL, DC)
            Funciones_Pic.motEnableDis(1, 1)
            ServoStopMotor(1, AMP_ENABLE Or MOTOR_OFF)
            ServoStopMotor(1, AMP_ENABLE Or STOP_SMOOTH)
            ServoResetPos(1)
        End If

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each sp As String In My.Computer.Ports.SerialPortNames
            ComboBox1.Items.Add(sp)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        COMPort = ComboBox1.SelectedItem.ToString & ":"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Home()

    End Sub
End Class
