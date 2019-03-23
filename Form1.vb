Imports System.Runtime.InteropServices
Public Class Form1
    Dim Original, Current As New DEVMODE With {.dmSize = CShort(Marshal.SizeOf(GetType(DEVMODE)))}
    Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As IntPtr) As IntPtr
    Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As IntPtr) As IntPtr
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Long
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFilename As String) As Boolean
#Region "DisplaySettings API"
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure DEVMODE
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public dmDeviceName As String
        Public dmSpecVersion As Short
        Public dmDriverVersion As Short
        Public dmSize As Short
        Public dmDriverExtra As Short
        Public dmFields As Integer
        Public dmOrientation As Short
        Public dmPaperSize As Short
        Public dmPaperLength As Short
        Public dmPaperWidth As Short
        Public dmScale As Short
        Public dmCopies As Short
        Public dmDefaultSource As Short
        Public dmPrintQuality As Short
        Public dmColor As Short
        Public dmDuplex As Short
        Public dmYResolution As Short
        Public dmTTOption As Short
        Public dmCollate As Short
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> _
        Public dmFormName As String
        Public dmLogPixels As Short
        Public dmBitsPerPel As Integer
        Public dmPelsWidth As Integer
        Public dmPelsHeight As Integer
        Public dmDisplayFlags As Integer
        Public dmDisplayFrequency As Integer
    End Structure
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function ChangeDisplaySettings(<[In]()> ByRef lpDevMode As DEVMODE, ByVal dwFlags As Integer) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
    Private Shared Function EnumDisplaySettings(ByVal lpszDeviceName As String, ByVal iModeNum As Integer, ByRef lpDevMode As DEVMODE) As Boolean
    End Function
#End Region
    Private Sub SettingsChanged() '配置项改变事件
        WritePrivateProfileString("Settings", "Bound", ComboBox1.SelectedIndex, Environment.CurrentDirectory & "\MineBound.ini")
        WritePrivateProfileString("Settings", "Refresh", CheckBox1.Checked, Environment.CurrentDirectory & "\MineBound.ini")
        WritePrivateProfileString("Settings", "Auto", CheckBox2.Checked, Environment.CurrentDirectory & "\MineBound.ini")
    End Sub
    Private Sub ActivationChanged() '激活状态改变事件
        EnumDisplaySettings(My.Computer.Screen.DeviceName, 0, Current)
        With Current
            .dmPelsWidth = My.Computer.Screen.Bounds.Width
            .dmPelsHeight = My.Computer.Screen.Bounds.Height
            .dmBitsPerPel = My.Computer.Screen.BitsPerPixel
            Label1.Text = "分辨率：" & .dmPelsWidth & "x" & .dmPelsHeight
            Label2.Text = "宽高比：" & IIf(.dmPelsHeight / .dmPelsWidth * 16 = 12, "4:3", "16:" & .dmPelsHeight / .dmPelsWidth * 16) ' y/x * 16
            Label3.Text = "刷新率：" & .dmDisplayFrequency & "Hz"
            Label4.Text = "色深度：" & .dmBitsPerPel & "Bit"
        End With
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Icon = My.Resources.BnB_MineSweeper
        If System.IO.File.Exists("MineBound.ini") Then
            Dim tString As New System.Text.StringBuilder
            GetPrivateProfileString("Settings", "Bound", -1, tString, 255, Environment.CurrentDirectory & "\MineBound.ini")
            ComboBox1.SelectedIndex = tString.ToString
            GetPrivateProfileString("Settings", "Refresh", False, tString, 255, Environment.CurrentDirectory & "\MineBound.ini")
            CheckBox1.Checked = Boolean.Parse(tString.ToString)
            GetPrivateProfileString("Settings", "Auto", False, tString, 255, Environment.CurrentDirectory & "\MineBound.ini")
            CheckBox2.Checked = Boolean.Parse(tString.ToString)
        End If '读取配置文件
        AddHandler CheckBox1.CheckedChanged, AddressOf SettingsChanged
        AddHandler CheckBox2.CheckedChanged, AddressOf SettingsChanged
        AddHandler ComboBox1.SelectedIndexChanged, AddressOf SettingsChanged
        AddHandler Activated, AddressOf ActivationChanged
        AddHandler Deactivate, AddressOf ActivationChanged
        '绑定事件处理函数
    End Sub
    Private Sub Form1_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Original = Current '置原始环境数据
        If CheckBox2.Checked Then
            Button1_Click(Nothing, Nothing)
        End If '自动应用分辨率
        If System.IO.File.Exists("ms_arbiter.exe") AndAlso FindWindowA("TMain", "Minesweeper Arbiter ").ToInt32 = 0 Then
            Process.Start("ms_arbiter.exe")
        End If '启动扫雷
    End Sub
    Private Sub Form1_Closed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        ChangeDisplaySettings(Original, 0) '还原
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If CheckBox1.Checked Then
            Dim tProcId As New IntPtr
            GetWindowThreadProcessId(FindWindowA("Shell_TrayWnd", Nothing), tProcId)
            Process.GetProcessById(tProcId).Kill()
        End If '重启资源管理器
        With Current
            Select Case ComboBox1.SelectedIndex
                Case 0
                    .dmPelsWidth = 1920
                    .dmPelsHeight = 1080
                Case 1
                    .dmPelsWidth = 1600
                    .dmPelsHeight = 900
                Case 2
                    .dmPelsWidth = 1280
                    .dmPelsHeight = 720
                Case 3
                    .dmPelsWidth = 1152
                    .dmPelsHeight = 864
                Case 4
                    .dmPelsWidth = 1024
                    .dmPelsHeight = 768
                Case 5
                    .dmPelsWidth = 800
                    .dmPelsHeight = 600
                Case 6
                    .dmPelsWidth = 640
                    .dmPelsHeight = 480
            End Select '变更边界参数
        End With
        ChangeDisplaySettings(Current, 0) '变更屏幕边界
        SetForegroundWindow(FindWindowA("TMain", "Minesweeper Arbiter ")) '呼出扫雷
    End Sub
End Class
