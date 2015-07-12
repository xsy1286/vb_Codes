Option Strict Off
Option Explicit On
Class Form1
    Inherits System.Windows.Forms.Form
    Private Structure POINTAPI
        Dim x As Integer
        Dim y As Integer
    End Structure
    'UPGRADE_WARNING: �ṹ POINTAPI ����Ҫ����ʹ���������Ϊ�� Declare ����еĲ������ݡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"��
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpoint As POINTAPI) As Integer

    'UPGRADE_NOTE: pv �������� pv_Renamed�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"��
    Dim p As POINTAPI
    Dim pv_Renamed As POINTAPI
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Const HWND_TOPMOST As Short = -1
    Const HWND_NOTOPMOST As Short = -2
    Const SWP_NOMOVE As Short = &H2S
    Const SWP_NOSIZE As Short = &H1S
    Const SWP_SHOWWINDOW As Short = &H40S
    Const SWP_NOOWNERZORDER As Short = &H200S '  Don't do owner Z ordering
    Dim top1 As Boolean
    Dim d As Short
    Dim t As Integer

    'UPGRADE_WARNING: ��ʼ������ʱ���ܼ����¼� Check1.CheckStateChanged�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"��
    Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged

        If top1 = False Then
            SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOOWNERZORDER) : top1 = True
        Else : SetWindowPos(Me.Handle.ToInt32, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOOWNERZORDER) : top1 = False
        End If
    End Sub

    Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        d = 0
        Timer1.Enabled = False
        t = 0
        Top = VB6.TwipsToPixelsY(False)
        Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
    End Sub

    Private Sub Form1_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y
        If Check1.CheckState = 1 Then
            d = 1
            Timer1.Interval = 5 : t = 0
            'Timer1.Enabled = True

            'UPGRADE_WARNING: δ�ܽ������� pa ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
            Call GetCursorPos(p)
        End If

    End Sub

    Private Sub Form1_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = eventArgs.X
        Dim y As Single = eventArgs.Y
        d = 0
        'UPGRADE_WARNING: ��ʱ������ Timer1.Interval ��ֵ����Ϊ 0�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"��
        Timer1.Enabled = False
        'UPGRADE_WARNING: ��ʱ������ Timer1.Interval ��ֵ����Ϊ 0�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"��

    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick

        t = t + 1
        'UPGRADE_WARNING: δ�ܽ������� p ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
        'UPGRADE_WARNING: δ�ܽ������� pv_Renamed ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
        pv_Renamed = p
        'UPGRADE_WARNING: δ�ܽ������� p ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
        GetCursorPos(p)
        'UPGRADE_WARNING: δ�ܽ������� p.x ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
        Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Left) + pv_Renamed.x - p.x)
        'UPGRADE_WARNING: δ�ܽ������� p.y ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
        Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Top) + pv_Renamed.y - p.y)

    End Sub
End Class