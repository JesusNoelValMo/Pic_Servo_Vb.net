Module Funciones_Pic
    '
    'NMC Communications functions
    '
    Declare Function NmcInit Lib "nmclib04" (ByVal portname As String, ByVal baudrate As Integer) As Integer
    Declare Sub NmcShutdown Lib "nmclib04" ()
    Declare Function NmcReadStatus Lib "nmclib04" (ByVal addr As Byte, ByVal statusitems As Byte) As Boolean
    Declare Function NmcNoOp Lib "nmclib04" (ByVal addr As Byte) As Boolean
    Declare Function NmcGetStat Lib "nmclib04" (ByVal addr As Byte) As Byte
    Public Const SEND_POS = &H1         '4 bytes data
    Public Const AMP_ENABLE = &H1       '1 = raise amp enable output, 0 = lower amp enable
    Public Const MOTOR_OFF = &H2        ' set to turn motor off
    Public Const STOP_ABRUPT = &H4      ' set to stop motor immediately
    Public Const STOP_SMOOTH = &H8      ' set to decellerate motor smoothly
    Public Const LOAD_POS = &H1         ' +4 bytes
    Public Const LOAD_VEL = &H2         ' +4 bytes
    Public Const LOAD_ACC = &H4         ' +4 bytes
    Public Const ENABLE_SERVO = &H10    ' 1 = servo mode, 0 = PWM mode
    Public Const START_NOW = &H80       ' 1 = start now, 0 = wait for START_MOVE command
    Public Const ON_LIMIT1 = &H1        ' home on change in limit 1
    Public Const ON_LIMIT2 = &H2        ' home on change in limit 2
    Public Const HOME_MOTOR_OFF = &H4   ' turn motor off when homed
    Public Const ON_INDEX = &H8           ' home on change in index
    Public Const HOME_STOP_ABRUPT = &H10  ' stop abruptly when homed
    Public Const MOVE_DONE = &H1        ' set when move done (trap. pos mode), when goal
    Public Const HOME_IN_PROG = &H80    ' set while searching for home, cleared when home found
    Public Const PATH_MODE = &H40       ' path mode is enabled (v.5)
    Declare Function ServoSetGain Lib "nmclib04" (ByVal addr As Byte, ByVal kp As Short, ByVal kd As Short, ByVal ki As Short, ByVal il As Short, ByVal ol As Byte, ByVal cl As Byte, ByVal el As Short, ByVal sr As Byte, ByVal dc As Byte) As Boolean
    Declare Function ServoResetPos Lib "nmclib04" (ByVal addr As Byte) As Boolean
    Declare Function ServoSetPos Lib "nmclib04" (ByVal addr As Byte, ByVal pos As Integer) As Boolean
    Declare Function ServoStopMotor Lib "nmclib04" (ByVal addr As Byte, ByVal mode As Byte) As Boolean
    Declare Function ServoSetHoming Lib "nmclib04" (ByVal addr As Byte, ByVal mode As Byte) As Boolean
    Declare Function ServoLoadTraj Lib "nmclib04" (ByVal addr As Byte, ByVal mode As Byte, ByVal pos As Integer, ByVal vel As Integer, ByVal acc As Integer, ByVal pwm As Byte) As Boolean
    Declare Function ServoGetPos Lib "nmclib04" (ByVal addr As Byte) As Integer
    Public Const SEND_INPUTS = &H1      ' 2 bytes data
    Public Const SEND_PERROR = &H40     '2 bytes
    Declare Function IoInBitVal Lib "nmclib04" (ByVal addr As Byte, ByVal bitnum As Byte) As Boolean
    Declare Function IoSetOutBit Lib "nmclib04" (ByVal addr As Byte, ByVal bitnum As Integer) As Boolean
    Declare Function IoClrOutBit Lib "nmclib04" (ByVal addr As Byte, ByVal bitnum As Integer) As Boolean
    Declare Function IoBitDirIn Lib "nmclib04" (ByVal addr As Byte, ByVal bitnum As Integer) As Boolean
    Declare Function IoBitDirOut Lib "nmclib04" (ByVal addr As Byte, ByVal bitnum As Integer) As Boolean
    ' Declare Function ServoAddPathPoints Lib "nmclib04" (ByVal addr As Byte, ByVal npoints As Integer, ByVal path() As Long, ByVal freq As Integer) As Boolean

    '****** Stepper Module Definitions ******
    '(note: some defined constants which are repeats of constants for the 
    ' PIC-SERVO are shown for your reference, but are commented out)
    Declare Function NmcSynchOutput Lib "nmclib04" (ByVal groupaddr As Byte, ByVal leaderaddr As Byte) As Boolean
    Declare Function NmcSetGroupAddr Lib "nmclib04" (ByVal addr As Byte, ByVal groupaddr As Byte, ByVal leader As Boolean) As Boolean



    Declare Function ServoAddPathpoints Lib "nmclib04" (ByVal addr As Byte, ByVal npoints As Integer, ByVal path() As Long, ByVal freq As Integer) As Boolean
    Declare Function ServoInitPath Lib "nmclib04" (ByVal addr As Byte) As Boolean
    Declare Function ServoStartPathMode Lib "nmclib04" (ByVal groupaddr As Byte, ByVal groupleader As Byte) As Boolean
    Public Const AMP_DISABLE As Integer = 0


    Public Function setGains(ByVal motor As Integer, ByVal kp As Integer, ByVal kd As Integer, ByVal ki As Integer, ByVal EL As Integer, ByVal DC As Integer)
        ServoSetGain(motor, kp, kd, ki, 0, 127, 0, EL, 3, DC)
        '   ServoSetGain(motor, kp, kd, ki, 0, 255, 0, EL, 1, DC)
    End Function
    Public Function motEnableDis(ByVal Motor As Integer, ByVal EnDis As Integer)
        If EnDis = 1 Then
            ServoStopMotor(Motor, AMP_ENABLE And MOTOR_OFF)
            ServoStopMotor(Motor, AMP_ENABLE And STOP_ABRUPT)
            ServoResetPos(Motor)
        Else
            ServoStopMotor(Motor, AMP_DISABLE And MOTOR_OFF)
            ServoStopMotor(Motor, AMP_DISABLE And STOP_ABRUPT)
            ServoResetPos(Motor)
        End If
    End Function
    Public Function Home()
        Dim statbyte(3) As Int32
        Dim mode As Byte
        Dim pos, vel, acc As Integer
        Dim pwm As Integer
        Dim homing_mode As Byte
        Dim a As Boolean

        Dim still_homing As Byte


        homing_mode = 0

        homing_mode = (ON_INDEX Or HOME_STOP_ABRUPT)
        mode = (LOAD_POS Or LOAD_VEL Or LOAD_ACC Or ENABLE_SERVO Or START_NOW)
        pos = -10000
        vel = 5000

        acc = 50
        pwm = 0
        addr = 1


        ServoSetHoming(addr, homing_mode)
        ServoLoadTraj(1, mode, pos, vel, acc, 0)


        Do


            NmcNoOp(addr)               'NoOp command to read current module status

            still_homing = (NmcGetStat(addr) And HOME_IN_PROG)

        Loop While still_homing
        addr = 2
        ServoSetHoming(addr, homing_mode)
        ServoLoadTraj(addr, LOAD_POS Or LOAD_VEL Or LOAD_ACC Or ENABLE_SERVO Or START_NOW, pos, vel, acc, pwm)
        Do

            NmcNoOp(addr)               'NoOp command to read current module status

            still_homing = (NmcGetStat(addr) And HOME_IN_PROG)

        Loop While still_homing
        addr = 3
        ServoSetHoming(addr, homing_mode)
        ServoLoadTraj(addr, LOAD_POS Or LOAD_VEL Or LOAD_ACC Or ENABLE_SERVO Or START_NOW, pos, vel, acc, pwm)
        Do

            NmcNoOp(addr)               'NoOp command to read current module status

            still_homing = (NmcGetStat(addr) And HOME_IN_PROG)

        Loop While still_homing
        NmcReadStatus(1, SEND_POS)  'Read the current position data from the cotnroller
        NmcReadStatus(2, SEND_POS)
        NmcReadStatus(3, SEND_POS)




    End Function
End Module
