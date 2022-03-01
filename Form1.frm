VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmKnxScsGate 
   BackColor       =   &H00000000&
   Caption         =   "KnxScsGate (TCP)"
   ClientHeight    =   9120
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   ScaleHeight     =   589.826
   ScaleMode       =   0  'User
   ScaleWidth      =   732.901
   Begin VB.CheckBox CheckTapPct 
      BackColor       =   &H80000012&
      Caption         =   "tapparelle %"
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   8280
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   13680
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame tcpFrame 
      Caption         =   "TCP - IP address"
      Height          =   615
      Left            =   4800
      TabIndex        =   11
      Top             =   7680
      Width           =   1815
      Begin VB.TextBox tcpText 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "192.168.2.180"
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   11640
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   6969
   End
   Begin VB.CommandButton PrintLog 
      Caption         =   "Print log"
      Height          =   495
      Left            =   10920
      TabIndex        =   10
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Print 
      Caption         =   "Print grid"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   7680
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Griglia1 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   12938
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   0
      BorderStyle     =   0
   End
   Begin VB.CommandButton Closeb 
      Caption         =   "Close channel"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Openb 
      Caption         =   "Open channel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton SerMonitor 
      Caption         =   "serial monitor"
      Height          =   495
      Left            =   9720
      TabIndex        =   5
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton query 
      Caption         =   "query firmware"
      Height          =   495
      Left            =   8640
      TabIndex        =   4
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   12120
      TabIndex        =   0
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton firmware 
      Caption         =   "new firmware"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   12720
      Top             =   7920
   End
   Begin RichTextLib.RichTextBox Rich 
      Height          =   7440
      Left            =   9660
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   13123
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   12120
      TabIndex        =   1
      Top             =   7680
      Width           =   2055
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10800
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   0   'False
      InBufferSize    =   4096
      BaudRate        =   19200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape shpCurr 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Menu Comm 
      Caption         =   "Communication"
      Begin VB.Menu Comm_1 
         Caption         =   "COM1"
      End
      Begin VB.Menu Comm_2 
         Caption         =   "COM2"
      End
      Begin VB.Menu Comm_3 
         Caption         =   "COM3"
      End
      Begin VB.Menu Comm_4 
         Caption         =   "COM4"
      End
      Begin VB.Menu Tcp_connection 
         Caption         =   "TCP"
      End
   End
   Begin VB.Menu FILE 
      Caption         =   "File"
      Begin VB.Menu LoadDataMenu 
         Caption         =   "Load data"
      End
      Begin VB.Menu SaveDataMenu 
         Caption         =   "Save data"
      End
      Begin VB.Menu DownloadMenu 
         Caption         =   "Download data"
      End
      Begin VB.Menu UploadMenu 
         Caption         =   "Upload data"
      End
   End
   Begin VB.Menu SYSTEM 
      Caption         =   "System"
      Begin VB.Menu Speed80Menu 
         Caption         =   "80Mhz"
      End
      Begin VB.Menu Speed160Menu 
         Caption         =   "160Mhz"
      End
   End
End
Attribute VB_Name = "frmKnxScsGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Const HOSTPORT_UDP = "52056"
Const HOSTPORT_TCP = "5045"
Const LOCALPORT = "6969"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public SerBuffer As String
Public tcpBufferIn As String
Public tcpBufferLen As Integer
Public tcpIpAddress As String
Public sFirmware As String
Public fwTimeout As Integer
Public tcpTimeout As Integer

Dim conntype As Integer
Dim tcpmode As Integer
Dim tcpcommand As String
Dim bauds As Long
Dim inPtr As Long
Dim ticks As Long
Dim iCom As Integer      ' comm serial port - 99=TCP
Dim sCom As String
Dim lngStatus As Long
Dim strError  As String
Dim devNumber

Dim sFlash As String
Dim maxflash As Long
Dim sAdr, sBlock, sChk As String
Dim smFLASH, iAck, myTimer, nrDevices As Integer
Dim dFile, dAdr, dAdrNxt, dSize, dFileLog As Long

Dim sMachine As Integer
Dim row As Integer
Dim sResp As String
Dim resp  As Integer
Dim retry As Integer
Dim hByte As Long
Dim lByte As Long
Dim nRepeat As Integer
Dim kByte As Long
Dim iMode As Byte  '  0=normale      1=serial monitor       2=gate
Dim bitVal(8) As Byte
Dim bNewCell As Boolean
Dim logcomment As Byte
Dim filelog As Integer

Dim gate As Byte   '  1=scs   2=konnex vimar   3=konnex true

Dim Telegramma(128, 7)

Private Const TLGR_ON = 0
Private Const TLGR_OFF = 1
Private Const TLGR_UP = 2
Private Const TLGR_DOWN = 3
Private Const TLGR_STOP = 4
Private Const TLGR_MORE = 5
Private Const TLGR_LESS = 6

Private Const MAINCAPTION As String = "KnxScsGate V6.0 (TCP)"
Private Const APPNAME As String = "KnxScsGate"
Private Const GR_LINE As Byte = 0
Private Const GR_ADDRESS As Byte = 1
Private Const GR_SUB As Byte = 2
Private Const GR_TIPO As Byte = 3
Private Const GR_DESCRI As Byte = 4
Private Const GR_COMMAND As Byte = 5
Private Const GR_STATUS  As Byte = 6
Private Const GR_SW1 As Byte = 7
Private Const GR_SW2 As Byte = 8
Private Const GR_SW3 As Byte = 9
Private Const GR_SW4 As Byte = 10
Private Const GR_MAXP As Byte = 11
Private Const GR_SWN As Byte = 90
                
Private Const GR_BACK_SWON = vbGreen
Private Const GR_BACK_SWOFF = vbGreen
Private Const GR_BACK_SWUP = vbGreen
Private Const GR_BACK_SWDOWN = vbGreen
Private Const GR_BACK_SWSTOP = vbGreen
Private Const GR_BACK_SWMORE = vbGreen
Private Const GR_BACK_SWLESS = vbGreen

Private Const GR_BACK_ON = vbYellow
Private Const GR_BACK_OFF = vbWhite

Private Const GR_CMD_SWON = "ON  "
Private Const GR_CMD_SWOFF = "OFF "
Private Const GR_CMD_SWUP = "UP  "
Private Const GR_CMD_SWDOWN = "DOWN"
Private Const GR_CMD_SWSTOP = "STOP"
Private Const GR_CMD_SWMORE = "MORE"
Private Const GR_CMD_SWLESS = "LESS"


Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type CharRange
  cpMin As Long ' First character of range (0 for start of doc)
  cpMax As Long ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hdc As Long ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As Rect ' Region of the DC to draw to (in twips)
  rcPage As Rect ' Region of the entire DC (page size) (in twips)
  chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
(ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
    

Private Sub cmdAbout_Click()
    MsgBox "Versione 4.0" & vbCr & "Made by G_Pagani" & vbCr & "Copyright (C) 2014" & vbCr, vbOK + vbInformation
End Sub

Private Sub cmdExit_Click()
    End
End Sub
Private Sub response_display()
    Dim iC As Integer
    Dim sHex 'As String
    
    sHex = ""
    sResp = ReadWait(16, 5)
    If sResp <> "" Then
        LogPrint "    response: " & Len(sResp) & " bytes : "
        For iC = 1 To Len(sResp)
            sHex = sHex & HexString(Asc(Mid(sResp, iC, 1)), 2)
        Next
        LogPrint (sHex)
        sHex = ""
    Else
        LogPrint "    response: <null>"
    End If

End Sub
Private Function write_display(sWrite)
    Dim iC As Integer
    Dim sHex 'As String
    
'    iC = Asc(Left(sWrite, 1)) + 1
'    sWrite = Left(sWrite + String(16, Chr(0)), iC)
    
    sWrite = Chr(Len(sWrite) + 2) + Chr(10) + Chr(7) + sWrite
    
    sHex = ""
    For iC = 1 To Len(sWrite)
        sHex = sHex & HexString(Asc(Mid(sWrite, iC, 1)), 2)
    Next
    LogPrint (sHex)
    WriteBuf (sWrite)
End Function
Private Sub firmware_Click()
    Dim iC, iD, iRc
    Dim sE 'As String
    Dim startFirmDir As String
    
    startFirmDir = GetSetting(APPNAME, "InitValues", "FirmDir")

    CommonDialog1.FileName = ""

    CommonDialog1.InitDir = startFirmDir
    CommonDialog1.DialogTitle = "Open  FLASH   bin file"
    CommonDialog1.DefaultExt = "hex"
    CommonDialog1.Filter = "*.hex"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
         sE = CommonDialog1.FileName
         
         startFirmDir = sE
         iC = Len(sE)
         While iC > 0 And Mid(startFirmDir, iC, 1) <> "\"
            iC = iC - 1
         Wend
         If iC > 2 Then
            startFirmDir = Left(startFirmDir, iC - 1)
         Else
            startFirmDir = Left(startFirmDir, 3)
         End If
         SaveSetting APPNAME, "InitValues", "FirmDir", startFirmDir
         
         smFLASH = 0  ' initial operations
         HexFlashFileRead (CommonDialog1.FileName)
         CommonDialog1.FileName = "temp.bin"

        frmKnxScsGate.Refresh
    
        resp = TryConnect(bauds)
        
        If resp = 0 Then
            LogPrint "send firmware update request"
            Flush
            fwTimeout = 0
            BinFlashFileRead (CommonDialog1.FileName)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close filelog
End Sub

Private Sub Griglia1_Click()
    Dim ptr, command, tClick
     
'    If gate = 1 Then    '  1=scs
'    If gate = 2 Then    '  2=konnex vimar

    ptr = Griglia1.row
    command = ""
    ' Griglia1.row   Griglia1.col è la cella su cui hai cliccato
    If Griglia1.row > 0 And (Griglia1.col >= GR_SW1 And Griglia1.col <= GR_SW4) And iMode = 2 Then
        tClick = RTrim(Griglia1.TextMatrix(Griglia1.row, Griglia1.col))
        
        If gate = 2 Then  '  1=konnex
            If tClick = "OFF" Then
                 command = Telegramma(ptr, TLGR_OFF)
                 Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
            End If
            If tClick = "ON" Then
                 command = Telegramma(ptr, TLGR_ON)
                 Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON)
            End If
            If tClick = "UP" Then
                 command = Telegramma(ptr, TLGR_UP)
                 Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWUP)
            End If
            If tClick = "DOWN" Then
                 command = Telegramma(ptr, TLGR_DOWN)
                 Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWDOWN)
            End If
            If tClick = "STOP" Then
                 command = Telegramma(ptr, TLGR_STOP)
                 Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWSTOP)
            End If
            
            If tClick = "MORE" Then
                 command = Telegramma(ptr, TLGR_MORE)
                 writeTlg (command)
                 sResp = ReadWait(1024, 9)
                 LogPrint (sResp)
                 command = Telegramma(ptr, TLGR_ON)
            End If
            If tClick = "LESS" Then
                 command = Telegramma(ptr, TLGR_LESS)
                 writeTlg (command)
                 sResp = ReadWait(1024, 9)
                 LogPrint (sResp)
                 command = Telegramma(ptr, TLGR_ON)
            End If
        End If

        If gate = 1 Then  '  1=scs
            If tClick = "OFF" Then
                 command = Telegramma(ptr, TLGR_OFF)
            End If
            If tClick = "ON" Then
                 command = Telegramma(ptr, TLGR_ON)
            End If
            
' scs 20161211
            If tClick = "MORE" Then
                 command = Telegramma(ptr, TLGR_MORE)
            End If
            If tClick = "LESS" Then
                 command = Telegramma(ptr, TLGR_LESS)
            End If
            
            If tClick = "UP" Then
                 command = Telegramma(ptr, TLGR_UP)
            End If
            If tClick = "DOWN" Then
                 command = Telegramma(ptr, TLGR_DOWN)
            End If
            If tClick = "STOP" Then
                 command = Telegramma(ptr, TLGR_STOP)
            End If
        End If
       
        If command <> "" Then
            writeTlg (command)
            sResp = ReadWait(1, 16)
            LogPrint (sResp)
        End If
    End If
End Sub
Function writeTlg(tlg As String)
    Dim lnw, lnwX
    lnw = Len(tlg)
    If lnw > 16 Then
        lnw = 16
        tlg = Left(tlg, 16)
    End If
    lnwX = Chr(lnw)
    WriteBuf ("@W" + lnwX + tlg)
    LogPrintContSer ("@W" + StringToHex(lnwX) + StringToHex(tlg))
End Function
Private Sub Griglia1_KeyPress(ikey As Integer)
    Dim sKey
    sKey = Chr(ikey)  ' 27: (esc) -> clear cell          8: (back)
    
    If Griglia1.row > 0 And Griglia1.col = GR_DESCRI Then
'       If bNewCell = True Then
'           If ikey <> 8 Then Griglia1.text = sKey
            bNewCell = False
'       Else
            If ikey = 27 Then ' esc key (clear)
                Griglia1.text = ""
            ElseIf ikey = 8 Then ' back key
                If Len(Griglia1.text) > 0 Then Griglia1.text = Left(Griglia1.text, Len(Griglia1.text) - 1)
            Else
                If Len(Griglia1.text) < 20 Then Griglia1.text = Griglia1.text & sKey
            End If
'       End If
    End If
    
  
    If Griglia1.row > 0 And (Griglia1.col = GR_ADDRESS Or Griglia1.col = GR_LINE) Then
'       If bNewCell = True Then
'            If ikey <> 8 Then Griglia1.text = sKey
            bNewCell = False
'        Else
            If ikey = 27 Then ' esc key (clear)
                Griglia1.text = ""
            ElseIf ikey = 8 Then ' back key
                If Len(Griglia1.text) > 0 Then Griglia1.text = Left(Griglia1.text, Len(Griglia1.text) - 1)
            ElseIf sKey >= "0" And sKey <= "9" Then
                If Len(Griglia1.text) < 2 Then Griglia1.text = Griglia1.text & sKey
            ElseIf sKey >= "a" And sKey <= "f" Then
                If Len(Griglia1.text) < 2 Then Griglia1.text = Griglia1.text & Chr(ikey - 32)
            ElseIf sKey >= "A" And sKey <= "F" Then
                If Len(Griglia1.text) < 2 Then Griglia1.text = Griglia1.text & sKey
            End If
 '       End If
    End If
    
    If Griglia1.row > 0 And Griglia1.col = GR_TIPO Then
 '      If bNewCell = True Then
 '          If ikey <> 8 Then Griglia1.text = sKey
            bNewCell = False
 '      Else
            If ikey = 27 Then ' esc key (clear)
                Griglia1.text = ""
            ElseIf ikey = 8 Then ' back key
                Griglia1.text = ""
            ElseIf sKey = "d" Or sKey = "D" Then
                Griglia1.text = "DIM"
            ElseIf sKey = "i" Or sKey = "I" Then
                Griglia1.text = "D+-"
            ElseIf sKey = "l" Or sKey = "L" Then ' Or sKey = "s" Then
                Griglia1.text = "LUCE"
            ElseIf sKey = "s" Or sKey = "S" Then    ' equivalente a TAP
                Griglia1.text = "STO"
            ElseIf sKey = "u" Or sKey = "U" Then
                Griglia1.text = "UPD"
            ElseIf sKey = "t" Or sKey = "T" Then
                Griglia1.text = "TAP"
            ElseIf sKey = "g" Or sKey = "G" Then
                Griglia1.text = "GEN"
            End If
'       End If
    End If
    
    If Griglia1.row > 0 And Griglia1.col = GR_MAXP Then
'       If bNewCell = True Then
'           If ikey <> 8 Then Griglia1.text = sKey
            bNewCell = False
'       Else
            If ikey = 27 Then ' esc key (clear)
                Griglia1.text = ""
            ElseIf ikey = 8 Then ' back key
                If Len(Griglia1.text) > 0 Then Griglia1.text = Left(Griglia1.text, Len(Griglia1.text) - 1)
            Else
                If Len(Griglia1.text) < 6 Then Griglia1.text = Griglia1.text & sKey
            End If
'       End If
    End If
End Sub

Private Sub Griglia1_Entercell()
    If (Griglia1.col = GR_ADDRESS Or Griglia1.col = GR_DESCRI) And Griglia1.row > 0 Then
        bNewCell = True
    End If
End Sub



Private Sub Openb_Click()
    Dim timout
    timout = 5
    If iMode = 0 Then
        Timer1.Enabled = False
        resp = TryConnect(bauds)
        
        If resp <> 0 Then
    '        Flush
            Disconnect
        Else
            iMode = 2
            
            WriteBuf ("@")
            WriteBuf (Chr(22))   '0x16
            LogPrintContSer ("@0x16")
            sResp = ReadWait(1024, timout)
            LogPrint (sResp)
            If sResp = "" Or Left(sResp, 1) <> "k" Then
                iMode = 0
            End If
            
            WriteBuf ("@MX")
            LogPrintContSer ("@MX")
            sResp = ReadWait(1024, timout)
            LogPrint (sResp)
            If sResp = "" Or Left(sResp, 1) <> "k" Then
                iMode = 0
            End If
            
            If iMode = 2 Then
                WriteBuf ("@F0")
                LogPrintContSer ("@F0")
                sResp = ReadWait(1024, timout)
                LogPrint (sResp)
                If sResp = "" Or Left(sResp, 1) <> "k" Then
                    iMode = 0
                End If
            End If
            
            If iMode = 2 Then
                WriteBuf ("@b")
                LogPrintContSer ("@b")
                sResp = ReadWait(1024, timout)
                LogPrint (sResp)
                If sResp = "" Or Left(sResp, 1) <> "k" Then
                    iMode = 0
                End If
            End If
            
            If iMode = 2 Then
                WriteBuf ("@Y0")
                LogPrintContSer ("@Y0")
                sResp = ReadWait(1024, timout)
                LogPrint (sResp)
                If sResp = "" Or Left(sResp, 1) <> "k" Then
                    iMode = 0
                End If
            End If
            
            If iMode = 2 Then
                WriteBuf ("@l")
                LogPrintContSer ("@l")
                sResp = ReadWait(1024, timout)
                LogPrint (sResp)
                If sResp = "" Or Left(sResp, 1) <> "k" Then
                    iMode = 0
                End If
            End If
            
            If iMode = 0 Then
                LogPrint "K.O. - communication failed..."
                Flush
                Disconnect
            Else
                firmware.Visible = False
                query.Visible = False
                SerMonitor.Visible = False
                Closeb.Visible = True
                Openb.Visible = False
                LogPrint "START connection"
                Timer1.Interval = 2
                Timer1.Enabled = True
            End If
         End If
    Else

    End If
    
End Sub
Private Sub Closeb_Click()
    If iMode = 2 Then
        Timer1.Enabled = False
        Disconnect
        iMode = 0
        firmware.Visible = True
        query.Visible = True
        SerMonitor.Visible = True
        Closeb.Visible = False
        Openb.Visible = True
        LogPrint "END connection"
    End If
End Sub



Private Sub query_Click()
        resp = TryConnect(bauds)
        
        If resp = 0 Then
            Flush
        End If
        Disconnect
End Sub
Function HexFlashFileRead(sFileName)
    Dim sZero, sDati, numLine, a
    Dim sLine, sCol, sTrk, sWork, sChex, sResp
    Dim iZ, iLen, iAdr, iExtAdr, iWork, iPtr, iEOF
    
    maxflash = 63488
    
    sFlash = String(maxflash, Chr(255))
   '
   ' Check if the file exists
   '
    dFile = FreeFile()
    Open sFileName For Input As dFile
    LogPrint "read & convert file..." & sFileName
    
    iExtAdr = 0
    numLine = 0
    Do
        Line Input #dFile, sLine
        numLine = numLine + 1
        sLine = UCase(sLine)
                
            ' destrutturazione del record
            ':LLAAAATTxxxxxxxxxxxxxxxxxxxxxxxxxCK
        sCol = Left(sLine, 1)
        iLen = DecByte(Mid(sLine, 2, 2))
        iAdr = DecWord(Mid(sLine, 4, 4))
        sTrk = Mid(sLine, 8, 2)
        sDati = Mid(sLine, 10, iLen * 2)

' trk 04: extended linear address record
        If sCol = ":" And sTrk = "04" And iLen > 0 Then
            iExtAdr = DecWord(Mid(sLine, 10, 4)) * 256 * 256
        End If
' trk 00: data record
        If sCol = ":" And sTrk = "00" And iLen > 0 Then
            iAdr = iAdr + iExtAdr
                        
' ============================== FLASH =======================================
            If iAdr >= 0 And iAdr < maxflash Then
                iPtr = iAdr + 1
                For iWork = 0 To iLen - 1
                'NON scambia LSB MSB (little endian) - LLMM LLMM LLMM
                     sChex = Mid(sDati, 1 + (iWork * 2), 2)
                     Mid(sFlash, iPtr, 1) = Chr(DecByte(sChex))
                     iPtr = iPtr + 1
                Next
            End If
                
        End If
    iEOF = iEOF + 1
' trk 01: fine file
    Loop Until EOF(dFile) Or (Len(sLine) < 3) Or (sCol <> ":") Or sTrk = "01" Or iEOF > 50000
    Close #dFile
    
    Open "temp.bin" For Binary Access Write As dFile
    Put dFile, 1, sFlash
    Close #dFile

    HexFlashFileRead = 1
End Function
Function BinFlashFileRead(sFileName)
    Dim i, iChk As Integer
    
    If smFLASH = 0 Then
        dAdr = 0
        dAdrNxt = 0

        dFile = FreeFile()
    
    'On Error GoTo BinNotExist
        Open sFileName For Binary As dFile
        dSize = LOF(dFile)
        
    'On Error GoTo BinSocketError
        LogPrint "Open socket..."
        LogPrint "Wait for ack..."
   
        WriteBuf ("@" + Chr(&H11))
        
        iAck = 0
        smFLASH = 1  ' wait for first ack
        BinFlashFileRead = 0
        
        sMachine = 20
        Timer1.Interval = 2
        Timer1.Enabled = True
        fwTimeout = 0
        Exit Function
    End If
    
    If smFLASH = 1 Then
        LogPrint "Ack received..."
        LogPrint "read & send bin file..." & sFileName
        LogPrint " "
        smFLASH = 2  ' ack received - send first block
    End If
    
    If smFLASH = 3 And Mid(sResp, 3, 1) > Chr(&HEF) Then  ' ack received
        smFLASH = 4 ' nack
    End If
    
    If smFLASH = 3 Then   ' ack received
        If dAdrNxt >= dSize Then
            Close dFile
            LogPrint " "
            LogPrint "End..."
'            Winsock.SendData "@FLE"    ' FINE --------------------@FLE------>
'            Call ShowData(">>@FLE", "")
            
            smFLASH = 0  ' wait for first ack
            sAdr = ""
            sMachine = 0
            Timer1.Enabled = False
' invia il comando 0x80 - end flash
'           write_firmware_init
'           write_firmware (Chr(&H80)) 'block length,   source device, pdu format, data=0x80(end)
'           write_firmware_end
            WriteBuf Chr(&H80)
'           Flush
            Disconnect

            BinFlashFileRead = 0
            Exit Function
        Else
            LogPrintCont ("k")
            smFLASH = 2
        End If
    End If
    
    If smFLASH = 4 Then   ' Nack received
         LogPrintCont ("r")
'        Winsock.SendData (sAdr + sBlock + sChk)  ' REINVIA lo stesso blocco ----->
        smFLASH = 3  ' block sended - wait for ack
        
        smFLASH = 20 ' resend same block
        BinFlashFileRead = 0
        Exit Function
    End If
    
    If smFLASH = 2 Then
        sBlock = ReadBin(64)
        smFLASH = 20
        fwTimeout = 0
    End If
    
    If smFLASH = 20 Then
        iChk = 0
        For i = 1 To 64
            iChk = iChk + Asc(Mid(sBlock, i, 1))
        Next
'        If sBlock <> sFF Then
        If iChk <> 16320 Then  ' significa tutti 0xFF
            sAdr = BinString2(dAdr)
            sChk = BinString2(iChk)
            LogPrintCont (".")
'            Winsock.SendData (sAdr + sBlock + sChk)    ' INVIO di un BLOCCO ------>
            
            smFLASH = 3  ' block sended - wait for ack
            BinFlashFileRead = 0
            
            write_firmware_init
' invia il comando 0x01 - new flash block binary mode
            write_firmware (Chr(1) + sAdr + Chr(64)) 'block length,   source device, pdu format, data=0x01(newblock), addressL, addressH, dataLength
' invia 8 blocchi da 8 bytes - flash binary data
            write_firmware (Mid(sBlock, 1, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 9, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 17, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 25, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 33, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 41, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 49, 8)) 'block length,   source device, pdu format, binary data 1-8
            write_firmware (Mid(sBlock, 57, 8)) 'block length,   source device, pdu format, binary data 1-8
    'sResp = ReadWait(8, 5)
' invia il comando 0x02 - end flash block binary mode
             write_firmware (Chr(2) + sChk) 'REALE block length,   source device, pdu format, data=0x02(endblock), checkL, checkH
'            write_firmware (Chr(3) + sChk) 'TEST  block length,   source device, pdu format, data=0x02(endblock), checkL, checkH
'            SlowDown (200)
            write_firmware_end
        Else
            smFLASH = 2  ' all 0xFF - block NOT sended - wait for loop
            BinFlashFileRead = 1  ' loop request
        End If
        Exit Function
    End If
    
    BinFlashFileRead = 0
    End Function
Function ReadBin(lgth As Long)
    Dim sChar As String
    
    sChar = Space(lgth)  ' space(lof(1))
    Get dFile, , sChar
    dAdr = dAdrNxt
    dAdrNxt = dAdrNxt + lgth
    ReadBin = sChar
End Function
Public Sub SlowDown(MilliSeconds As Long)
Dim lngTickStore As Long
    lngTickStore = GetTickCount()
    Do While lngTickStore + MilliSeconds > GetTickCount()
    DoEvents
    Loop
End Sub

Private Sub Rich_KeyPress(KeyAscii As Integer)
    Dim sChar As String
    
    If iMode = 1 Then
'        LogPrint (">" + Chr(KeyAscii))
        sChar = Chr(KeyAscii)
        
        If sChar = "*" Then
            If logcomment = 0 Then
                logcomment = 1
                LogPrintCont ("***__")
            Else
                logcomment = 0
                LogPrintCont ("__***")
            End If
        End If
            
        Print #filelog, sChar;
        
        If logcomment = 0 Then
            WriteBuf (sChar)
        End If
'       sResp = ReadWait(1024, 5)
'       LogPrint ("response: " & sResp)
    End If
End Sub

Private Sub SerMonitor_Click()
    If iMode = 0 Then
        Timer1.Enabled = False
        resp = TryConnect(bauds)
        
        If resp <> 0 Then
            Flush
            Disconnect
        Else
            iMode = 1
            firmware.Visible = False
            query.Visible = False
            Openb.Visible = False
            LogPrint "START serial ascii monitor mode"
            
            WriteBuf ("@MA")
            LogPrintContSer ("@MA")
            sResp = ReadWait(1024, 5)
            LogPrint (sResp)

            Timer1.Interval = 2
            Timer1.Enabled = True
        End If
    Else
        Timer1.Enabled = False
        Disconnect
        iMode = 0
        firmware.Visible = True
        query.Visible = True
        Openb.Visible = True
        LogPrint "END serial monitor mode"

    End If
End Sub
Private Sub OpenTCP()
    If (conntype = 1) Or (conntype = 11) Or (conntype = 12) Then
        Winsock1.Close
        conntype = 0
    End If
    
    If (conntype = 0) Then
      Winsock1.Protocol = sckTCPProtocol
      Winsock1.RemoteHost = tcpText.text
      Winsock1.RemotePort = HOSTPORT_TCP
      conntype = 1  'tcp request
      Winsock1.Connect
      LogPrint ("Open TCP socket on " + tcpText.text + " port " + HOSTPORT_TCP)
      DoEvents
    ElseIf (conntype = 2) Then
      Call Winsock1_Connect
    End If
End Sub
Private Sub DownloadMenu_Click()
    tcpmode = -1  ' download
    Call OpenTCP
End Sub

Private Sub Speed80Menu_Click()
    tcpmode = -2 ' direct command
    tcpcommand = "#setup {""frequency"":""80""}"
    Call OpenTCP
End Sub

Private Sub Speed160Menu_Click()
    tcpmode = -2 ' direct command
    tcpcommand = "#setup {""frequency"":""160""}"
    Call OpenTCP
End Sub

Private Sub UploadMenu_Click()
    Dim conferma
    
    conferma = MsgBox("Rimpiazzo TUTTI i dispositivi censiti?", vbYesNo)
    
    If conferma = 6 Then
        tcpmode = 1  ' upload
        tcpcommand = "#setup {""frequency"":""160""}"
        Call OpenTCP
    End If
End Sub
Private Sub Winsock1_Close()
    tcpText.BackColor = &H80000005
    conntype = 0
    tcpmode = 0
    If (conntype = 1) Or (conntype = 2) Then
      LogPrint ("TCP closed")
    End If
    If (conntype = 11) Or (conntype = 12) Then
      LogPrint ("TCP closed")
    End If
    Winsock1.Close
End Sub
Private Sub Winsock1_Connect()
Dim max_row, max_col, R As Integer
Dim device, sectorline, devtype, descr, maxp As String
    
    tcpText.BackColor = &HFF00&
    
    If (conntype = 1) Then  'tcp request
        conntype = 2
        LogPrint ("TCP opened")
    End If
    
    If (conntype = 2) Then  'tcp request
      If (tcpmode = -2) Then
        tcpmode = 0
        LogPrint (tcpcommand)
        Winsock1.SendData tcpcommand
      End If
    
'      If (tcpmode = -1) Then ' download
'        LogPrint ("> download...")
    '   Winsock1.SendData "GET /status?pippo=ABCDHTTP/1.1 " + vbCrLf + "Host: " + tcpText.text + vbCrLf
'        Winsock1.SendData "#getdevall {}"
'      End If

      If (tcpmode = -1) Then ' download
        LogPrint ("> download...")
    '   Winsock1.SendData "GET /status?pippo=ABCDHTTP/1.1 " + vbCrLf + "Host: " + tcpText.text + vbCrLf
        devNumber = 1
'        descr = "#getdevice {""devnum"":""" + Str$(devNumber) + """}"
        Winsock1.SendData "#getdevice {""devnum"":""" + Str$(devNumber) + """}"
      End If
    
      If (tcpmode > 0) Then ' upload
'        If (CheckTapPct.Value = 1) Then
'            LogPrint ("#putdevice {""coverpct"":""true"",""devclear"":""true""}")
'            Winsock1.SendData "#putdevice {""coverpct"":""true"",""devclear"":""true""}"
'        Else
            LogPrint ("#putdevice {""coverpct"":""false"",""devclear"":""true""}")
            Winsock1.SendData "#putdevice {""coverpct"":""false"",""devclear"":""true""}"
'        End If
      End If
    End If
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim msg As String
Dim ln As Integer
    
    If (conntype = 2) Then '2->tcp request
'------------------------------------------------------------------
  '    LogPrint ("< received data...")
      Winsock1.GetData msg
'      LogPrint (msg)                   ' SOLO PER DEBUG
    
      If (tcpmode = 0) Then '0->nothing to do
'         Winsock1.Close
'         conntype = 0
      End If
    
      If (tcpmode = -1) Then '-1-> download
        If Left$(msg, 4) = "#eof" Then
            tcpmode = 0
'           Winsock1.Close
'           conntype = 0
            LogPrint (msg)
            Griglia1.Rows = Griglia1.Rows + 10
        Else
            Call getNextRow(msg)
        End If
      End If
    
      If (tcpmode > 0) Then '1-255-> upload
        If Left$(msg, 3) = "#ok" Then
            Call putNextRow
        Else
            LogPrint (msg)
        End If
'        If Left$(msg, 4) = "#err" Then
'            Call putNextRow
'        End If
      End If

      If (tcpmode = -10) Then  '-10-> dequeue
            ln = Len(msg)
            tcpBufferIn = tcpBufferIn & msg
            tcpBufferLen = tcpBufferLen + ln
      End If
      
'-----------------------------------------------------------------------
    End If
    
End Sub
Private Sub getNextRow(ByRef msg As String)
Dim sectorline, address, device, devtype, maxpos, descr As String
Dim ptr As Integer
        
' download - "msg" row received

    ' "device":"0B31","type":"1","maxp":"0","descr":"luce trentauno"
    
    device = tcpJarg(msg, """device""")
'   LogPrint (">> get row device " + device)
    LogPrint (">> get row " + msg)  ' <<--------------------- DEBUG ------------------
    If (device <> "") Then
        devtype = tcpJarg(msg, """type""")
        maxpos = tcpJarg(msg, """maxp""")
        descr = tcpJarg(msg, """descr""")
        If Len(device) >= 4 Then
            sectorline = Left$(device, 2)
            address = Mid$(device, 3)
        Else
            sectorline = "12"
            address = device
        End If
        ptr = SearchTabHex(sectorline, address)
        Griglia1.row = ptr
        
        If (devtype = "1") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "LUCE"
        ElseIf (devtype = "3") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM"
        ElseIf (devtype = "4") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM"
        ElseIf (devtype = "24") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "D+-"
        ElseIf (devtype = "8") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "TAP"
'            Griglia1.TextMatrix(ptr, GR_TIPO) = "STO"
        ElseIf (devtype = "9") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "TAP"
'            Griglia1.TextMatrix(ptr, GR_TIPO) = "STO"
            CheckTapPct.Value = 1
        ElseIf (devtype = "18") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "UPD"
        ElseIf (devtype = "19") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "UPD"
            CheckTapPct.Value = 1
        ElseIf (devtype = "11") Then
            Griglia1.TextMatrix(ptr, GR_TIPO) = "GEN"
        Else
            Griglia1.TextMatrix(ptr, GR_TIPO) = " ? "
        End If
        
        Griglia1.TextMatrix(ptr, GR_DESCRI) = descr
        Griglia1.TextMatrix(ptr, GR_MAXP) = maxpos
        devNumber = devNumber + 1
        
        Griglia1.row = ptr
        Griglia1.col = GR_ADDRESS
        Griglia1.CellBackColor = vbYellow
        
        Winsock1.SendData "#getdevice {""devnum"":""" + Str$(devNumber) + """}"
    End If
End Sub
Private Sub putNextRow()
Dim max_row, max_col, R As Integer
Dim device, sectorline, devtype, devtypeDes, descr, maxp As String
    
    If (tcpmode > 0) Then ' upload
        max_row = Griglia1.Rows - 1
        Do
            devtype = ""
            sectorline = Griglia1.TextMatrix(tcpmode, GR_LINE)
            device = Griglia1.TextMatrix(tcpmode, GR_ADDRESS)
            descr = Griglia1.TextMatrix(tcpmode, GR_DESCRI)
            maxp = Griglia1.TextMatrix(tcpmode, GR_MAXP)
            devtypeDes = Griglia1.TextMatrix(tcpmode, GR_TIPO)
            If (devtypeDes = "LUCE") Then
                devtype = "1"
            ElseIf (devtypeDes = "GEN") Then
                devtype = "11"
            ElseIf (devtypeDes = "DIM") Then
                If gate = 1 Then    '  1=scs
                    devtype = "3"
                Else
                    devtype = "4"
                End If
            ElseIf (devtypeDes = "D+-") Then
                devtype = "24"
            
            ElseIf (devtypeDes = "TAP") Then
                If CheckTapPct.Value = 1 Then
                    devtype = "9"
                Else
                    devtype = "8"
                End If
            ElseIf (devtypeDes = "STO") Then
                If CheckTapPct.Value = 1 Then
                    devtype = "9"
                Else
                    devtype = "8"
                End If
            ElseIf (devtypeDes = "UPD") Then
                If CheckTapPct.Value = 1 Then
                    devtype = "19"
                Else
                    devtype = "18"
                End If
            End If
            tcpmode = tcpmode + 1
        Loop While (tcpmode <= max_row) And ((devtype = "") Or (device = "") Or (sectorline = "" And gate = 2))
        
        If ((devtype <> "") And (device <> "") And (sectorline <> "" Or gate = 1)) Then
        
           If gate = 1 Then    '  1=scs
             LogPrint (">> send row device " + device)
             Winsock1.SendData "#putdevice {""device"":""" + device + """,""type"":""" + devtype + """,""maxp"":""" + maxp + """,""descr"":""" + descr + """}"
             LogPrint "#putdevice {""device"":""" + device + """,""type"":""" + devtype + """,""maxp"":""" + maxp + """,""descr"":""" + descr + """}" ' <------------------------- DEBUG
             Griglia1.row = tcpmode - 1
             Griglia1.col = GR_ADDRESS
             Griglia1.CellBackColor = vbGreen
           ElseIf gate = 2 Then    '  2=konnex vimar
             Winsock1.SendData "#putdevice {""device"":""" + sectorline + device + """,""type"":""" + devtype + """,""maxp"":""" + maxp + """,""descr"":""" + descr + """}"
             LogPrint (">> send row device " + sectorline + device)
             Griglia1.row = tcpmode - 1
             Griglia1.col = GR_ADDRESS
             Griglia1.CellBackColor = vbGreen
           Else
             LogPrint ("SCS/KNX unknown - please make query function and retry")
             tcpmode = 0
           End If

        End If
'        tcpmode = tcpmode + 1
        
        
        If (tcpmode > max_row) Then
'           tcpmode = 0
            tcpmode = -2 ' direct command
            tcpcommand = "#setup {""commit"":""true""}"
            Call OpenTCP
        End If
        
    End If
    
End Sub
Private Function tcpJarg(buffer As String, argument As String) As String
    Dim valore As String
    Dim p1, p2, p3, p4 As Integer
    valore = ""
    p1 = InStr(1, buffer, argument)
    If (p1 > 0) Then
        p2 = InStr(p1 + 1, buffer, ":")
        If (p2 > 0) Then
            p3 = InStr(p2 + 1, buffer, Chr(34)) ' prima "
            If (p3 > 0) Then
                p4 = InStr(p3 + 1, buffer, Chr(34)) ' successiva "
                If (p4 > 0) And ((p4 - p3) > 1) Then
                    valore = Mid$(buffer, p3 + 1, p4 - p3 - 1)
                End If
            End If
        End If
    End If
    tcpJarg = valore
End Function
Private Sub Timer1_Timer()
    Dim lMsg As Integer
    ticks = ticks + 1
    
    If sMachine = 20 Then
        ' FIRMWARE update in corso !!!!
        ' riceve   0xLL  destination   pduformat    data...
        sResp = ReadNoWait(1)
        If sResp <> "" Then
            fwTimeout = 0
''''            LogPrintContHex (sResp)
            lMsg = Asc(sResp)
            sResp = ReadWait(lMsg, 5)
''''            LogPrintHex (sResp)
            Do
                Loop While (BinFlashFileRead("") = 1)
        Else
            fwTimeout = fwTimeout + 1
            If fwTimeout > 20 Then
                fwTimeout = 0
                sResp = Chr(0) + " " + Chr(&HF5) + "  "
                Do
                    Loop While (BinFlashFileRead("") = 1)
            End If
        End If
    Else
        tcpTimeout = tcpTimeout + 1
        If tcpTimeout > 400 Then
                tcpTimeout = 0
                KeepAlive
        End If
    End If

    If iMode = 0 Then
        sResp = ReadNoWait(16)
    End If

    If iMode = 1 Then
        sResp = ReadNoWait(16)
       
        If sResp <> "" Then
            LogPrintContSer (sResp)
        End If
    End If
    
    If iMode = 2 Then
        sResp = ReadNoWait(1)
        If sResp <> "" Then
            lMsg = Asc(sResp)
            sResp = ReadWait(lMsg, 5)
        End If
'
        
        If sResp <> "" Then
            LogPrintHex (sResp)
            If gate = 2 Then DecodeKNX (sResp) '  2=konnex vimar
            If gate = 1 Then DecodeSCS (sResp) '  1=scs
        End If
    End If
    
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim lMsg As Integer
    
    ticks = ticks + 1
    
    DequeueTCP
    
End Sub

Private Function DecodeSCS(sResp)
    Dim prefix, origin, destination, network, command, check, suffix As String
    Dim ptr As Integer
    
    prefix = Mid(sResp, 1, 1)
    If prefix = Chr(&HA5) Then Exit Function
    If Len(sResp) < 7 Then Exit Function
'    If Len(sResp) > 7 Then exit function
    If Len(sResp) > 7 Then sResp = Left(sResp, 7)
    
    destination = Mid(sResp, 2, 1)
    origin = Mid(sResp, 3, 1)
    network = Mid(sResp, 4, 1)
    command = Mid(sResp, 5, 1)
    check = Mid(sResp, 6, 1)
    suffix = Mid(sResp, 7, 1)
        
    If prefix = Chr(&HA8) And suffix = Chr(&HA3) And origin = Chr(&H0) Then

' è un comando
        ptr = SearchTab(network, destination)
        Griglia1.row = ptr
        Griglia1.TextMatrix(ptr, GR_COMMAND) = StringToHex(command)
        If command = Chr(0) Then
            Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON)
            If Griglia1.TextMatrix(ptr, GR_TIPO) = "" Then
                Griglia1.TextMatrix(ptr, GR_TIPO) = "LUCE"
            End If
            Telegramma(ptr, TLGR_ON) = sResp
            If (Telegramma(ptr, TLGR_OFF) = "") Then
                Telegramma(ptr, TLGR_OFF) = Left(sResp, 4) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1) + suffix
            End If
        End If
        If command = Chr(1) Then
            Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
            If Griglia1.TextMatrix(ptr, GR_TIPO) = "" Then
                Griglia1.TextMatrix(ptr, GR_TIPO) = "LUCE"
            End If
            Telegramma(ptr, TLGR_OFF) = sResp
            If (Telegramma(ptr, TLGR_ON) = "") Then
                Telegramma(ptr, TLGR_ON) = Left(sResp, 4) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1) + suffix
            End If
        End If
        If command = Chr(3) Then
'            Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
'            If Griglia1.TextMatrix(ptr, GR_TIPO) = "" Then
                Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM"
                Griglia1.TextMatrix(ptr, GR_STATUS) = "MORE"
                Call setGrid(Griglia1, GR_SW3, GR_BACK_SWMORE, True, GR_CMD_SWMORE)
                Call setGrid(Griglia1, GR_SW4, GR_BACK_SWLESS, True, GR_CMD_SWLESS)
            Telegramma(ptr, TLGR_MORE) = sResp
            If (Telegramma(ptr, TLGR_LESS) = "") Then
                Telegramma(ptr, TLGR_LESS) = Left(sResp, 4) + Chr(Asc(command) Xor &H7) + Chr(Asc(check) Xor &H7) + suffix
            End If

'            End If
        End If
        If command = Chr(4) Then
'            Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
'            If Griglia1.TextMatrix(ptr, GR_TIPO) = "" Then
                Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM"
                Griglia1.TextMatrix(ptr, GR_STATUS) = "LESS"
                Call setGrid(Griglia1, GR_SW3, GR_BACK_SWMORE, True, GR_CMD_SWMORE)
                Call setGrid(Griglia1, GR_SW4, GR_BACK_SWLESS, True, GR_CMD_SWLESS)
            Telegramma(ptr, TLGR_LESS) = sResp
            If (Telegramma(ptr, TLGR_MORE) = "") Then
                Telegramma(ptr, TLGR_MORE) = Left(sResp, 4) + Chr(Asc(command) Xor &H7) + Chr(Asc(check) Xor &H7) + suffix
            End If
'            End If
        End If
        Call setGrid(Griglia1, GR_SW1, GR_BACK_SWON, True, GR_CMD_SWON)
        Call setGrid(Griglia1, GR_SW2, GR_BACK_SWOFF, True, GR_CMD_SWOFF)
    End If
    
    
    If prefix = Chr(&HA8) And suffix = Chr(&HA3) And destination = Chr(&HB8) Then
' è uno stato
        ptr = SearchTab(network, origin)
        Griglia1.row = ptr
        If command = Chr(0) Then
            Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON)
        End If
        If command = Chr(1) Then
            Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
        End If
        If (Asc(command) And &HD) = &HD Then
            Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON + " " + Chr(&H30 + Asc(command) \ 16))
        End If

'        Griglia1.TextMatrix(ptr, GR_TIPO) = "LUCE"
'        Call setGrid(Griglia1, GR_SW1, GR_BACK_SWON, True, GR_CMD_SWON)
'        Call setGrid(Griglia1, GR_SW2, GR_BACK_SWOFF, True, GR_CMD_SWOFF)
    End If
End Function
Private Function DecodeKNX(sResp)   '  ----->>>>>>>>>>>>>>> ADATTARE A STO UPD TAP <<<<<<<<<<<<<<<<<
    Dim prefix, origin, destination, dsub, network, command, check As String
    Dim ptr, addressMain, addressSub, cmdlen As Integer
    
    prefix = Mid(sResp, 1, 1)
    If prefix = Chr(&HCC) Then Exit Function
    If prefix = Chr(&H5F) Then Exit Function
    If Len(sResp) < 8 Then Exit Function
    
    ' B4 10 0E 0B 1D E1 00 81 23
    
    destination = Mid(sResp, 5, 1)
    addressMain = Asc(destination)
    addressSub = Asc(destination) - Int(Asc(destination) / 2) * 2
    If addressSub = 1 Then   ' dispari - normale ( pulsante premuto breve )
        addressSub = 0
    Else                    ' pari (sub! - pulsante premuto lungo)
        addressSub = 1
        addressMain = addressMain - 1
    End If
    If addressMain > 0 Then
        destination = Chr(addressMain)
        Else
        destination = Chr(0)
    End If
    dsub = HexString(addressSub, 2)
    
    origin = Mid(sResp, 3, 1)
    network = Mid(sResp, 4, 1)
    cmdlen = Asc(Mid(sResp, 6, 1)) And &HF
    command = Mid(sResp, 8, cmdlen)
    check = Mid(sResp, 8 + cmdlen, 1)
 
    If prefix = Chr(&HB4) Or prefix = Chr(&HB8) Or prefix = Chr(&HBC) Then
' è un comando konnex
        ptr = SearchTab(network, destination)
        Griglia1.row = ptr
        If addressSub > 0 Then Griglia1.TextMatrix(ptr, GR_SUB) = dsub
        dsub = Griglia1.TextMatrix(ptr, GR_SUB)
        Griglia1.TextMatrix(ptr, GR_COMMAND) = StringToHex(command)
        
        If addressSub = 0 Then ' sub corrente 0 (tasto premuto BREVE)
            If Griglia1.TextMatrix(ptr, GR_TIPO) = "TAP" Or Griglia1.TextMatrix(ptr, GR_TIPO) = "TAP%" Then ' tapparella ********************************
                If command = Chr(&H80) Then
'                   Griglia1.TextMatrix(ptr, GR_STATUS) = "STOP"
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWSTOP)
                    Telegramma(ptr, TLGR_STOP) = sResp
                End If
                If command = Chr(&H81) Then
'                    Griglia1.TextMatrix(ptr, GR_STATUS) = "STOP"
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWSTOP)
                    Telegramma(ptr, TLGR_STOP) = sResp
                End If
            Else
            If Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM" Then ' dimmer ************************************
                If command = Chr(&H80) Then
'                    Griglia1.TextMatrix(ptr, GR_STATUS) = "STOP"
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
                    Telegramma(ptr, TLGR_OFF) = sResp
                    If (Telegramma(ptr, TLGR_ON) = "") Then
                        Telegramma(ptr, TLGR_ON) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
                    End If
                End If
                If command = Chr(&H81) Then
'                    Griglia1.TextMatrix(ptr, GR_STATUS) = "STOP"
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON)
                    Telegramma(ptr, TLGR_ON) = sResp
                    If (Telegramma(ptr, TLGR_OFF) = "") Then
                        Telegramma(ptr, TLGR_OFF) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
                    End If
                End If
            Else                                              ' luce **************************************
                If command = Chr(&H80) Then
'                   Griglia1.TextMatrix(ptr, GR_STATUS) = "OFF"
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_OFF, False, GR_CMD_SWOFF)
                    Telegramma(ptr, TLGR_OFF) = sResp
                    If (Telegramma(ptr, TLGR_ON) = "") Then
                        Telegramma(ptr, TLGR_ON) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
                    End If
                End If
                If command = Chr(&H81) Then
'                   Griglia1.TextMatrix(ptr, GR_STATUS) = "ON"
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON)
                    Telegramma(ptr, TLGR_ON) = sResp
                    If (Telegramma(ptr, TLGR_OFF) = "") Then
                        Telegramma(ptr, TLGR_OFF) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
                    End If
                End If
                Griglia1.TextMatrix(ptr, GR_TIPO) = "LUCE"
                Call setGrid(Griglia1, GR_SW1, GR_BACK_SWON, True, GR_CMD_SWON)
                Call setGrid(Griglia1, GR_SW2, GR_BACK_SWOFF, True, GR_CMD_SWOFF)
            End If
            End If
        End If
        
        If addressSub > 0 Then                 ' tasto premuto a lungo
            If command = Chr(&H80) Then
                If Griglia1.TextMatrix(ptr, GR_TIPO) <> "DIM" Then
                    Griglia1.TextMatrix(ptr, GR_STATUS) = "UP"
                    Griglia1.TextMatrix(ptr, GR_TIPO) = "TAP"
                    Call setGrid(Griglia1, GR_SW1, GR_BACK_SWUP, True, GR_CMD_SWUP)
                    Call setGrid(Griglia1, GR_SW2, GR_BACK_SWDOWN, True, GR_CMD_SWDOWN)
                    Call setGrid(Griglia1, GR_SW3, GR_BACK_SWSTOP, True, GR_CMD_SWSTOP)
                    Telegramma(ptr, TLGR_UP) = sResp
                    If (Telegramma(ptr, TLGR_DOWN) = "") Then
                        Telegramma(ptr, TLGR_DOWN) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
                    End If
                Else
'                    Griglia1.TextMatrix(ptr, GR_STATUS) = "LESS"
'                    Telegramma(ptr, TLGR_LESS) = sResp
                    Call setGrid(Griglia1, GR_STATUS, GR_BACK_ON, False, GR_CMD_SWON)
'                    Telegramma(ptr, TLGR_ON) = sResp
'                    If (Telegramma(ptr, TLGR_LESS) = "") Then
'                        Telegramma(ptr, TLGR_LESS) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
'                    End If
                End If
            End If
            If command = Chr(&H81) Then
                If Griglia1.TextMatrix(ptr, GR_TIPO) <> "DIM" Then
                    Griglia1.TextMatrix(ptr, GR_STATUS) = "DOWN"
                    Griglia1.TextMatrix(ptr, GR_TIPO) = "TAP"
                    Call setGrid(Griglia1, GR_SW1, GR_BACK_SWUP, True, GR_CMD_SWUP)
                    Call setGrid(Griglia1, GR_SW2, GR_BACK_SWDOWN, True, GR_CMD_SWDOWN)
                    Call setGrid(Griglia1, GR_SW3, GR_BACK_SWSTOP, True, GR_CMD_SWSTOP)
                    Telegramma(ptr, TLGR_DOWN) = sResp
                    If (Telegramma(ptr, TLGR_UP) = "") Then
                        Telegramma(ptr, TLGR_UP) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
                    End If
                Else
                    Griglia1.TextMatrix(ptr, GR_STATUS) = "LESS"
                    Telegramma(ptr, TLGR_LESS) = sResp
'                    If (Telegramma(ptr, TLGR_ON) = "") Then
'                        Telegramma(ptr, TLGR_ON) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
'                    End If
                End If
            End If
            If command = Chr(&H88) Then
                Griglia1.TextMatrix(ptr, GR_STATUS) = "ON"
                Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM"
                Call setGrid(Griglia1, GR_SW1, GR_BACK_SWON, True, GR_CMD_SWON)
                Call setGrid(Griglia1, GR_SW2, GR_BACK_SWOFF, True, GR_CMD_SWOFF)
                Call setGrid(Griglia1, GR_SW3, GR_BACK_SWMORE, True, GR_CMD_SWMORE)
                Call setGrid(Griglia1, GR_SW4, GR_BACK_SWLESS, True, GR_CMD_SWLESS)
'                Telegramma(ptr, TLGR_ON) = sResp
'                If (Telegramma(ptr, TLGR_MORE) = "") Then
'                    Telegramma(ptr, TLGR_MORE) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
'                End If
            End If
            If command = Chr(&H89) Then
                Griglia1.TextMatrix(ptr, GR_STATUS) = "MORE"
                Griglia1.TextMatrix(ptr, GR_TIPO) = "DIM"
                Call setGrid(Griglia1, GR_SW1, GR_BACK_SWON, True, GR_CMD_SWON)
                Call setGrid(Griglia1, GR_SW2, GR_BACK_SWOFF, True, GR_CMD_SWOFF)
                Call setGrid(Griglia1, GR_SW3, GR_BACK_SWMORE, True, GR_CMD_SWMORE)
                Call setGrid(Griglia1, GR_SW4, GR_BACK_SWLESS, True, GR_CMD_SWLESS)
                Telegramma(ptr, TLGR_MORE) = sResp
 '                   If (Telegramma(ptr, TLGR_ON) = "") Then
 '                       Telegramma(ptr, TLGR_ON) = Left(sResp, 7) + Chr(Asc(command) Xor 1) + Chr(Asc(check) Xor 1)
 '                   End If
            End If
        End If
    End If
End Function


Private Function SearchTab(line, address)
    Dim addressX, lineX

    lineX = StringToHex(line)
    addressX = StringToHex(address)
    SearchTab = SearchTabHex(lineX, addressX)
End Function
Private Function SearchTabHex(lineX, addressX)
    Dim iRow, found, emptyRow
    emptyRow = 0
    lineX = RTrim(LTrim(lineX))
    addressX = RTrim(LTrim(addressX))
    
    iRow = 0
    found = 0
    While (found = 0)
        Griglia1.row = iRow
           '      2=konnex
        If Griglia1.TextMatrix(iRow, GR_ADDRESS) = "" And (emptyRow = 0) Then
            emptyRow = iRow
        End If
        If Griglia1.TextMatrix(iRow, GR_ADDRESS) = addressX And (gate = 1 Or Griglia1.TextMatrix(iRow, GR_LINE) = lineX) Then
            found = 1
            SearchTabHex = iRow
            Exit Function
        End If
        iRow = iRow + 1
        If iRow >= Griglia1.Rows Then
            If emptyRow = 0 Then
                Griglia1.Rows = iRow + 1
                Griglia1.row = iRow
                found = 1
                SearchTabHex = iRow
                Griglia1.TextMatrix(iRow, GR_LINE) = lineX
                Griglia1.TextMatrix(iRow, GR_ADDRESS) = addressX
            Else
                iRow = emptyRow
                Griglia1.row = iRow
                found = 1
                SearchTabHex = iRow
                Griglia1.TextMatrix(iRow, GR_LINE) = lineX
                Griglia1.TextMatrix(iRow, GR_ADDRESS) = addressX
            End If
            Exit Function
        End If
    Wend
    
End Function
Function setGrid(ByVal flxData As MSFlexGrid, col, back, bold, text)
    flxData.col = col
    flxData.CellBackColor = back
    flxData.CellFontBold = bold
    flxData.text = text
   ' flxData.CellPicture
End Function

Private Sub Form_Load()
         Dim t As Integer
         Dim x As Long
         Dim t1, t2
         
         frmKnxScsGate.Caption = MAINCAPTION
         bauds = 115200
    
         sMachine = 0
         iMode = 0
         conntype = 0
         tcpmode = 0
         logcomment = 0
         
         bitVal(1) = &H80
         bitVal(2) = &H40
         bitVal(3) = &H20
         bitVal(4) = &H10
         bitVal(5) = &H8
         bitVal(6) = &H4
         bitVal(7) = &H2
         bitVal(8) = &H1
         t = 0
    tcpFrame.Visible = False

    t1 = Now    ' 15/03/2021 15:40:02
    t2 = Mid(Now, 7, 4) + Mid(Now, 4, 2) + Mid(Now, 1, 2) + "_" + Mid(Now, 12, 2) + Mid(Now, 15, 2) + Mid(Now, 18, 2)
    filelog = FreeFile
    Open "knxscslog" + t2 + ".txt" For Output As filelog
    
    LogPrint (t2)
    
    iCom = 5
    sCom = GetSetting(APPNAME, "InitValues", "ComPort")
    Select Case sCom
        Case "1"
            Comm_1_Click
        Case "2"
            Comm_2_Click
        Case "3"
            Comm_3_Click
        Case "4"
            Comm_3_Click
        Case "99"
            Tcp_connection_Click
    End Select
    
    tcpText.text = GetSetting(APPNAME, "InitValues", "TCP_IP4")
    
    Griglia1.Visible = False
    Griglia1.row = 0
    Griglia1.FormatString = ">     |>     |>   |>     |<__                    |>   |>       |>     |>     |>     |>     |>      "
    Griglia1.Visible = True
    Griglia1.Rows = 1
    Griglia1.Clear
    Griglia1.Font = "Courier new"
    Griglia1.row = 0
    Griglia1.TextMatrix(0, GR_LINE) = "Line"
    Griglia1.TextMatrix(0, GR_ADDRESS) = "Adrs"
    Griglia1.TextMatrix(0, GR_SUB) = "Sub"
    Griglia1.TextMatrix(0, GR_TIPO) = "Tipo "
    Griglia1.TextMatrix(0, GR_DESCRI) = "Device name"
    Griglia1.TextMatrix(0, GR_COMMAND) = "last cmd"
    Griglia1.TextMatrix(0, GR_STATUS) = "Status"
    Griglia1.TextMatrix(0, GR_SW1) = "CLK "
    Griglia1.TextMatrix(0, GR_SW2) = "  COM"
    Griglia1.TextMatrix(0, GR_SW3) = "MANDS"
    Griglia1.TextMatrix(0, GR_SW4) = "    "
    Griglia1.TextMatrix(0, GR_MAXP) = "maxP"
'   Griglia1.BackColorBkg = vbRed

'    Call setGrid(Griglia1, GR_SW1, GR_BACK_SWON, True, GR_CMD_SWON)
    frmKnxScsGate.Width = 16500
    frmKnxScsGate.Height = 10000

    Griglia1.Refresh
End Sub

Private Sub Comm_1_Click()
    iCom = 1
    LogPrint ("communication on COM1:")
    Comm_1.Checked = True
    Comm_2.Checked = False
    Comm_3.Checked = False
    Comm_4.Checked = False
    Tcp_connection.Checked = False
    tcpFrame.Visible = False
    SaveSetting APPNAME, "InitValues", "ComPort", iCom
End Sub

Private Sub Comm_2_Click()
    iCom = 2
    LogPrint ("communication on COM2:")
    Comm_2.Checked = True
    Comm_1.Checked = False
    Comm_3.Checked = False
    Comm_4.Checked = False
    Tcp_connection.Checked = False
    tcpFrame.Visible = False

    SaveSetting APPNAME, "InitValues", "ComPort", iCom
End Sub

Private Sub Comm_3_Click()
    iCom = 3
    LogPrint ("communication on COM3:")
    Comm_3.Checked = True
    Comm_2.Checked = False
    Comm_1.Checked = False
    Comm_4.Checked = False
    Tcp_connection.Checked = False
    tcpFrame.Visible = False

    SaveSetting APPNAME, "InitValues", "ComPort", iCom
End Sub

Private Sub Comm_4_Click()
    iCom = 4
    LogPrint ("communication on COM4:")
    Comm_4.Checked = True
    Comm_2.Checked = False
    Comm_3.Checked = False
    Comm_1.Checked = False
    Tcp_connection.Checked = False
    tcpFrame.Visible = False

    SaveSetting APPNAME, "InitValues", "ComPort", iCom
End Sub
Private Sub Tcp_connection_Click()
    iCom = 99
    LogPrint ("communication on TCP port:" + HOSTPORT_TCP)
    Comm_4.Checked = False
    Comm_2.Checked = False
    Comm_3.Checked = False
    Comm_1.Checked = False
    Tcp_connection.Checked = True
    tcpFrame.Visible = True

    SaveSetting APPNAME, "InitValues", "ComPort", iCom
End Sub
Function BinString2(ThisNumber)
    Dim iHnumber, iLnumber As Long
    '
    ' Convert an integer to a BIN STRING (2 bytes LOW_ENDIAN)
    '
    iHnumber = ThisNumber \ 256
    iLnumber = ThisNumber - (iHnumber * 256)
    BinString2 = Chr(iLnumber) + Chr(iHnumber)
End Function

Function DecWord(String4c)
' converte un esadecimale espresso ascii (4 bytes string) in esadecimale puro (BIN STRING 2 byte)
    Dim Ret1
    Dim Ret2
    '
    Ret1 = DecByte(Left(String4c, 2))
    Ret2 = DecByte(Right(String4c, 2))

    DecWord = Ret2 + (Ret1 * 256)

End Function
Function DecByte(String2c)
' converte un esadecimale espresso ascii (2 bytes string) in esadecimale puro (BIN STRING 1 byte)
    Dim RetVal, hexval
    hexval = "0123456789ABCDEF"
    If Len(String2c) > 2 Then String2c = Left(String2c, 2)
    RetVal = 16 * (InStr(hexval, Left(String2c, 1)) - 1)

    DecByte = RetVal + (InStr(hexval, Right(String2c, 1)) - 1)

End Function

Function StringToDec(sBuffer)
    ' Convert a string (1 or more bytes LOW_ENDIAN) into decimal double int
    Dim sLine As String
    Dim iPointer, iZ, iMult As Double
    Dim iTot As Double
        
    sLine = ""
    iTot = 0
    iMult = 1
    iPointer = Len(sBuffer)
    If sBuffer <> "" Then
        Do
            iZ = Asc(Mid(sBuffer, iPointer, 1))
            iTot = iTot + iZ * iMult
            sLine = sLine & HexString(iZ, 2)
            iPointer = iPointer - 1
            iMult = iMult * 256
         Loop While iPointer > 0
     End If
    sLine = iTot
    StringToDec = sLine
End Function
            
Function StringToHex(sBuffer)
'  convert string BIN STRING  in ascii hex  HEX ASCII STRING
    Dim sLine
    Dim iPointer, iZ
    
    sLine = ""
    iPointer = 1
    If sBuffer <> "" Then
        Do
            iZ = Asc(Mid(sBuffer, iPointer, 1))
            sLine = sLine & HexString(iZ, 2) & " "
            iPointer = iPointer + 1
         Loop While iPointer <= Len(sBuffer)
    End If
    StringToHex = sLine
End Function


Function BinString(sBuffer)
'  convert string BIN STRING  in ascii hex  HEX ASCII STRING
    Dim iZ, iPointer
    iPointer = 1
   ' sWork = ""
    Do
        iZ = Asc(Mid(sBuffer, iPointer, 1))
        BinString = BinString & HexString(iZ, 2)
        iPointer = iPointer + 1
    Loop Until iPointer > Len(sBuffer)
End Function

Function HexString(ThisNumber, length)
    '
    ' Convert an integer to a ascii hex string of requested length padding with the desired number of zeros
    '
    Dim RetVal
    Dim CurLen
    RetVal = Hex(ThisNumber)
    CurLen = Len(RetVal)

    If CurLen < length Then
        RetVal = String(length - CurLen, "0") & RetVal
    End If

    HexString = RetVal
End Function
Function hTrim(ThisString)
    Do While Right(ThisString, 1) = Chr(0)
        ThisString = Left(ThisString, Len(ThisString) - 1)
    Loop 'Until Right(ThisString, 1) = Chr(0)
    hTrim = ThisString
End Function

Function xTrim(ThisString)
    Dim iX
    For iX = 1 To Len(ThisString)
        If Mid(ThisString, iX, 1) < Chr(20) Then
            Mid(ThisString, iX, 1) = " "
        End If
    Next
    xTrim = ThisString
End Function
Function xString(ThisString)
    Dim iX
    iX = 0
    Do
        iX = iX + 1
    Loop Until (iX > Len(ThisString)) Or (Mid(ThisString, iX, 1) = Chr(0))
    xString = Left(ThisString, iX - 1)
End Function

Private Sub PrintLog_Click()
    Printer.Font.Name = Griglia1.Font.Name
    Printer.Font.Size = Griglia1.Font.Size
'    Printer.ScaleTop = 800
'    Call Rich.SelPrint(Printer.hdc)
    Call PrintRTF(Rich, 800, 800, 800, 800)

End Sub
Private Sub PrintLog_Click1()
    Screen.MousePointer = vbHourglass
    Dim LineWidth As Long
    Rich.SelFontName = "Arial"
    Rich.SelFontSize = 10
    LineWidth = WYSIWYG_RTF(Rich, 1440, 1440) '1440 Twips=1 Inch
    PrintRTF Rich, 1440, 1440, 1440, 1440 ' 1440 Twips = 1 Inch
    Screen.MousePointer = vbDefault
End Sub



Private Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)

Dim LeftOffset As Long
Dim TopOffset As Long
Dim LeftMargin As Long
Dim TopMargin As Long
Dim RightMargin As Long
Dim BottomMargin As Long
Dim fr As FormatRange
Dim rcDrawTo As Rect
Dim rcPage As Rect
Dim TextLength As Long
Dim NextCharPosition As Long
Dim R As Long

' Start a print job to get a valid Printer.hDC
Printer.Print Space(1)
Printer.ScaleMode = vbTwips

' Get the offsett to the printable area on the page in twips
LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

' Calculate the Left, Top, Right, and Bottom margins
LeftMargin = LeftMarginWidth - LeftOffset
TopMargin = TopMarginHeight - TopOffset
RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset

BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

' Set printable area rect
rcPage.Left = 0
rcPage.Top = 0
rcPage.Right = Printer.ScaleWidth
rcPage.Bottom = Printer.ScaleHeight

' Set rect in which to print (relative to printable area)
rcDrawTo.Left = LeftMargin
rcDrawTo.Top = TopMargin
rcDrawTo.Right = RightMargin
rcDrawTo.Bottom = BottomMargin

' Set up the print instructions
fr.hdc = Printer.hdc ' Use the same DC for measuring and rendering
fr.hdcTarget = Printer.hdc ' Point at printer hDC
fr.rc = rcDrawTo ' Indicate the area on page to draw to
fr.rcPage = rcPage ' Indicate entire size of page
fr.chrg.cpMin = 0 ' Indicate start of text through
fr.chrg.cpMax = -1 ' end of the text

' Get length of text in RTF
TextLength = Len(RTF.text)

' Loop printing each page until done
Do
' Print the page by sending EM_FORMATRANGE message
NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
If NextCharPosition >= TextLength Then Exit Do 'If done then exit
fr.chrg.cpMin = NextCharPosition ' Starting position for next page
Printer.NewPage ' Move on to next page
Printer.Print Space(1) ' Re-initialize hDC
fr.hdc = Printer.hdc
fr.hdcTarget = Printer.hdc
Loop

' Commit the print job
Printer.EndDoc

' Allow the RTF to free up memory
R = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
End Sub

Private Function WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long) As Long

Dim LeftOffset As Long
Dim LeftMargin As Long
Dim RightMargin As Long
Dim LineWidth As Long
Dim PrinterhDC As Long
Dim R As Long

' Start a print job to initialize printer object
Printer.Print Space(1)
Printer.ScaleMode = vbTwips

' Get the offset to the printable area on the page in twips
LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)

' Calculate the Left, and Right margins
LeftMargin = LeftMarginWidth - LeftOffset
RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset

' Calculate the line width
LineWidth = RightMargin - LeftMargin

' Create an hDC on the Printer pointed to by the Printer object
' This DC needs to remain for the RTF to keep up the WYSIWYG display
PrinterhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)

' Tell the RTF to base it's display off of the printer
' at the desired line width
R = SendMessage(RTF.hWnd, EM_SETTARGETDEVICE, PrinterhDC, ByVal LineWidth)

' Abort the temporary print job used to get printer info
Printer.KillDoc

WYSIWYG_RTF = LineWidth
End Function



Private Sub Print_Click()
    Call PrintFlexGrid(Printer, Griglia1, 800, 800)
    Printer.EndDoc
End Sub

Private Function PrintFlexGrid(ByVal ptr As Object, ByVal flxData As MSFlexGrid, ByVal xmin As Single, ByVal ymin As Single)
Const GAP = 60

Dim xmax As Single
Dim ymax As Single
Dim x As Single
Dim C As Integer
Dim R As Integer
    
    ptr.Font.Name = flxData.Font.Name
    ptr.Font.Size = flxData.Font.Size

    With flxData
        ' See how wide the whole thing is.
        xmax = xmin + GAP
        For C = 0 To .Cols - 1
            xmax = xmax + .ColWidth(C) + 2 * GAP
        Next C

        ' Print each row.
        ptr.CurrentY = ymin
        For R = 0 To .Rows - 1
            ' Draw a line above this row.
            If R > 0 Then ptr.Line (xmin, _
                ptr.CurrentY)-(xmax, ptr.CurrentY)
            ptr.CurrentY = ptr.CurrentY + GAP

            ' Print the entries on this row.
            x = xmin + GAP
            For C = 0 To .Cols - 1
                ptr.CurrentX = x
                ptr.Print BoundedText(ptr, .TextMatrix(R, _
                    C), .ColWidth(C));
                x = x + .ColWidth(C) + 2 * GAP
            Next C
            ptr.CurrentY = ptr.CurrentY + GAP

            ' Move to the next line.
            ptr.Print
        Next R
        ymax = ptr.CurrentY

        ' Draw a box around everything.
        ptr.Line (xmin, ymin)-(xmax, ymax), , B

        ' Draw lines between the columns.
        x = xmin
        For C = 0 To .Cols - 2
            x = x + .ColWidth(C) + 2 * GAP
            ptr.Line (x, ymin)-(x, ymax)
        Next C
    End With
End Function

' Truncate the string so it fits within the width.
Private Function BoundedText(ByVal ptr As Object, ByVal txt _
    As String, ByVal max_wid As Single) As String
    Do While ptr.TextWidth(txt) > max_wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    BoundedText = txt
End Function
Private Sub SaveDataMenu_Click()
    Dim iC, iD, iRc
    Dim sE 'As String
    Dim startDataDir As String
    
    startDataDir = GetSetting(APPNAME, "InitValues", "DataDir")

    CommonDialog1.FileName = ""

    CommonDialog1.InitDir = startDataDir
    CommonDialog1.DialogTitle = "Save  DATA file"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "*.txt"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
         sE = CommonDialog1.FileName
         
         startDataDir = sE
         iC = Len(sE)
         While iC > 0 And Mid(startDataDir, iC, 1) <> "\"
            iC = iC - 1
         Wend
         If iC > 2 Then
            startDataDir = Left(startDataDir, iC - 1)
         Else
            startDataDir = Left(startDataDir, 3)
         End If
         SaveSetting APPNAME, "InitValues", "DataDir", startDataDir
         
        Call SaveData(Griglia1, CommonDialog1.FileName)
    End If
End Sub
Private Sub LoadDataMenu_Click()
    Dim iC, iD, iRc
    Dim sE 'As String
    Dim startDataDir As String
    
    startDataDir = GetSetting(APPNAME, "InitValues", "DataDir")

    CommonDialog1.FileName = ""

    CommonDialog1.InitDir = startDataDir
    CommonDialog1.DialogTitle = "Load  DATA file"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "*.txt"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
         sE = CommonDialog1.FileName
         
         startDataDir = sE
         iC = Len(sE)
         While iC > 0 And Mid(startDataDir, iC, 1) <> "\"
            iC = iC - 1
         Wend
         If iC > 2 Then
            startDataDir = Left(startDataDir, iC - 1)
         Else
            startDataDir = Left(startDataDir, 3)
         End If
         SaveSetting APPNAME, "InitValues", "DataDir", startDataDir
         
        Call LoadData(Griglia1, CommonDialog1.FileName)
    End If
End Sub


' Save the FlexGrid data
Private Sub SaveData(ByVal flxData As MSFlexGrid, ByVal file_name As String)
Dim fnum As Integer
Dim max_row As Integer
Dim max_col As Integer
Dim R As Integer
Dim C As Integer


    fnum = FreeFile
    Open file_name For Output As fnum

    ' Save the maximum row and column.
    max_row = flxData.Rows - 1
    max_col = flxData.Cols - 1
    Write #fnum, max_row, max_col

    For R = 0 To max_row
        For C = 0 To max_col
            Write #fnum, flxData.TextMatrix(R, C);
        Next C
        Write #fnum,
    Next R

    Close fnum
End Sub


 
Private Sub LoadData(ByVal flxData As MSFlexGrid, ByVal file_name As String)
Dim fnum As Integer
Dim max_row As Integer
Dim max_col As Integer
Dim R As Integer
Dim C As Integer
Dim txt As String
Dim max_len As Single
Dim new_len As Single

    fnum = FreeFile
    Open file_name For Input As fnum

    ' Hide the control until it's loaded.
    flxData.Visible = False
    DoEvents

    ' Get the maximum row and column.
    Input #fnum, max_row, max_col

  '  flxData.FixedCols = 0
    flxData.Cols = max_col + 1
  '  flxData.FixedRows = 1
    flxData.Rows = max_row + 1

    ' Load the cell entries.
    For R = 0 To max_row
        For C = 0 To max_col
            Input #fnum, txt
            flxData.TextMatrix(R, C) = txt
        Next C
        ' Read the last blank entry.
        Input #fnum, txt
    Next R
    Close #fnum

    ' Display the control.
    flxData.Visible = True
End Sub
Private Sub tcpText_Change()
    SaveSetting APPNAME, "InitValues", "TCP_IP4", tcpText
End Sub
Function LogPrint(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = The_Data & vbCrLf
        .SelColor = vbBlack
    End With
    Rich.Refresh
  
    Print #filelog, The_Data
'    Write #filelog, The_Data & vbCrLf
End Function
Function LogPrintCont(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = The_Data
        .SelColor = vbBlack
    End With
    Rich.Refresh
    Print #filelog, The_Data;
'    Write #filelog, The_Data
End Function
Function LogPrintContSer(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = The_Data
        .SelColor = vbBlack
    End With
    Rich.Refresh
    Print #filelog, The_Data;
'    Write #filelog, The_Data
End Function
Function LogPrintHex(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = StringToHex(The_Data) & vbCrLf
        .SelColor = vbBlack
    End With
    Rich.Refresh
    Print #filelog, StringToHex(The_Data)
'    Write #filelog, StringToHex(The_Data) & vbCrLf
End Function
Function LogPrintContHex(ByVal The_Data As String, Optional ByVal ThisColor As Long = 0) As Boolean
Dim oldcolor
Dim start1

    With Rich
        start1 = Len(.text)
        .SelStart = start1
        oldcolor = .SelColor
        .SelColor = ThisColor
        .SelText = StringToHex(The_Data)
        .SelColor = vbBlack
    End With
    Rich.Refresh
    Print #filelog, StringToHex(The_Data);
'    Write #filelog, StringToHex(The_Data);
End Function

' ===================================================================================================================
'  COMMUNICATION SECTION - GENERIC
' ===================================================================================================================
Function TryConnect(iBaud)
    If (iCom = 99) Then
        TryConnect = TryConnectTCP(iBaud)
    Else
        TryConnect = TryConnectCOM(iBaud)
    End If
End Function

Private Sub Disconnect()
    If (iCom = 99) Then
        DisconnectTCP
    Else
        DisconnectCOM
    End If
End Sub
Function Write_synch(sString)
    Dim sResp, k, y
    WriteBuf sString
'    For k = 1 To 1000000: y = y + 1 / 7: Next
    sResp = ReadWait(1, 250)
    Select Case sResp
    Case "k"
       Write_synch = 1
    Case ""
       LogPrint "xxxGate non risponde al comando (" & Left(sString, 1) & ")... [" & sResp & "]"
       MsgBox ("xxxGate non risponde al comando (" & Left(sString, 1) & ")...")
       Write_synch = 0
    Case Else
       LogPrint "xxxGate non riconosce il comando (" & Left(sString, 1) & ")... [" & sResp & "]"
       MsgBox ("xxxGate non riconosce il comando (" & Left(sString, 1) & ")...")
       Write_synch = 0
    End Select
End Function

Function Wait_synch()
    Dim sResp
    sResp = ReadWait(1, 1)
    If (sResp = "" Or sResp <> "k") Then
       LogPrint "xxxGate non riceve ack comando precedente- rx: " & sResp
       MsgBox ("xxxGate non riceve ack comando precedente")
       Wait_synch = 0
    Else
       Wait_synch = 1
    End If
End Function
Private Function KeepAlive()
    If (iCom = 99) Then
'        WriteBufTCP ("@Keep_alive")
    End If
End Function
Private Function write_firmware_init()
    sFirmware = ""
End Function
Private Function write_firmware_end()
    If (iCom = 99) Then
        WriteBufTCP (sFirmware)
        sFirmware = ""
    End If
End Function
Private Function write_firmware(sWrite)
' invia i dati in una "busta" CAN-like
    If (iCom = 99) Then
        sFirmware = sFirmware + Chr(Len(sWrite) + 2) + Chr(10) + Chr(7) + sWrite
    Else
       WriteBufCOM (Chr(Len(sWrite) + 2) + Chr(10) + Chr(7) + sWrite)
    End If
End Function

Function WriteBuf(sString)
    If (iCom = 99) Then
        WriteBuf = WriteBufTCP(sString)
    Else
        WriteBuf = WriteBufCOM(sString)
    End If
End Function

Function ReadWait(length, timeout)
    If (iCom = 99) Then
        ReadWait = ReadWaitTCP(length, timeout)
    Else
        ReadWait = ReadWaitCOM(length, timeout)
    End If
End Function

Function ReadNoWait(length)
    If (iCom = 99) Then
        ReadNoWait = ReadNoWaitTCP(length)
    Else
        ReadNoWait = ReadNoWaitCOM(length)
    End If
End Function

Function Flush()
    If (iCom = 99) Then
        Flush = FlushTCP()
    Else
        Flush = FlushCOM()
    End If
End Function


' ===================================================================================================================
'  COMMUNICATION SECTION - USART
' ===================================================================================================================
Function TryConnectCOM(iBaud)
    Dim sResp
    LogPrint ("COM port initialize")
    On Error Resume Next        ' Abilito l'intercettazione degli errori
    ' Initialize Communications
    lngStatus = CommClose(iCom)
    lngStatus = CommOpen(iCom, "COM" & CStr(iCom), "baud=" & iBaud & " parity=N data=8 stop=1")
    If lngStatus <> 0 Then
    ' Handle error.
        lngStatus = CommGetError(strError)
        MsgBox "COM Error: " & strError
        TryConnectCOM = Err
        Exit Function
    End If
    On Error GoTo 0
    
    Flush
    LogPrint "TRYING communication"
    ' verifico se risponde a query come programmatore
    Flush
    WriteBuf ("@q")
    sResp = ReadWait(1024, 5)
    LogPrint ("response: " & sResp)
    'response_display
    If Len(sResp > 1) And Left(sResp, 2) = "qk" Then
'        bEcho = True
        sResp = Mid(sResp, 2)
    End If
    If sResp = "" Or Left(sResp, 1) <> "k" Then
        LogPrint "K.O. - communication failed..."
        Flush
        TryConnectCOM = 1 ' ko - retry
'        MSComm1.PortOpen = False    ' Chiudiamo la porta.
    Else
        If Len(sResp) > 1 Then
            LogPrint "OK!! - firmware version is " & Mid(sResp, 2)
            If Mid(sResp, 2, 3) = "KNX" Then
               gate = 2
               frmKnxScsGate.Caption = MAINCAPTION + " - KNX MODE"
            Else
               gate = 1
               frmKnxScsGate.Caption = MAINCAPTION + " - SCS MODE"
               Openb.Enabled = True
            End If
        Else
            LogPrint "OK!! "
        End If
        TryConnectCOM = 0 ' ok
    End If
End Function

Private Sub DisconnectCOM()
    Dim iC
    gate = 0
    frmKnxScsGate.Caption = MAINCAPTION
    On Error Resume Next    ' Abilito l'intercettazione degli errori
    lngStatus = CommClose(iCom)
    LogPrint "SERIAL CHANNEL CLOSED"
End Sub

Function WriteBufCOM(sString)
'    ShapeCom.FillColor = &HFF&
'    ShapeCom.Refresh
    Dim iP, iTim

    For iP = 1 To Len(sString)
        lngStatus = CommWrite(iCom, Mid(sString, iP, 1))
        'For iTim = 1 To 1999: Next
    Next
'    ShapeCom.FillColor = &HFF00&
'    ShapeCom.Refresh
'    If bEcho = True Then
'        Flush
'    End If
End Function

Function ReadWaitCOM(length, timeout)
    Dim tim
    Dim Chread As String
    Dim nrbytes
    Dim lngSize As Long
    ReadWaitCOM = ""
    Chread = " "
    tim = 0
    lngSize = 1
    Do
        lngStatus = CommRead(iCom, Chread, lngSize)
        If lngStatus <= 0 Then
            tim = tim + 1
        Else
            SerBuffer = SerBuffer & Chread
        End If
    Loop Until (Len(SerBuffer) >= length Or tim >= 10000 * timeout) ' 5 secondi timeout
    
    If Len(SerBuffer) > length Then
        ReadWaitCOM = Left(SerBuffer, length)
        SerBuffer = Mid(SerBuffer, length + 1)
    Else
        ReadWaitCOM = SerBuffer
        SerBuffer = ""
    End If
End Function

Function ReadNoWaitCOM(length)
    Dim tim
    Dim Chread As String
    Dim nrbytes
    Dim lngSize As Long
    ReadNoWaitCOM = ""
    tim = 0
    lngSize = 1
    Chread = " "
        lngStatus = CommRead(iCom, Chread, lngSize)
        If lngStatus <= 0 Then
            tim = tim + 1
        Else
            SerBuffer = SerBuffer & Chread
        End If
    
    If Len(SerBuffer) > length Then
        ReadNoWaitCOM = Left(SerBuffer, length)
        SerBuffer = Mid(SerBuffer, length + 1)
    Else
        ReadNoWaitCOM = SerBuffer
        SerBuffer = ""
    End If
End Function

Function FlushCOM()
    Dim iCol, Chread, SerBufferHex

    lngStatus = CommFlush(iCom)
    If Len(SerBuffer) > 0 Then
        LogPrint Len(SerBuffer) & " bytes flushed: "
        For iCol = 1 To Len(SerBuffer)
            SerBufferHex = SerBufferHex & HexString(Asc(Mid(SerBuffer, iCol, 1)), 2)
        Next
        LogPrint (SerBufferHex)
        SerBuffer = ""
    End If
End Function
Public Sub PauseConnect(Second As Double)
On Error Resume Next

Dim Dim_Dbl_AtTime As Double

Dim_Dbl_AtTime = Timer

Do While (Timer - Dim_Dbl_AtTime < Val(Second)) And (conntype <> 2)
  DoEvents
Loop

End Sub
' ===================================================================================================================
'  COMMUNICATION SECTION - TCP
' ===================================================================================================================
Function TryConnectTCP(tcpIpAddress)

    Dim sResp
    LogPrint ("TCP port initialize")
    On Error GoTo TryError
    
    tcpmode = -2 ' direct command
    tcpcommand = "#setup {""uart"":""tcp""}"
    Call OpenTCP
    Call PauseConnect(3)
    sResp = ReadWaitTCP(1024, 5)
    
    If (conntype = 2) Then
      tcpmode = -10 ' dequeue command
      
      WriteBufTCP ("@q")
      sResp = ReadWaitTCP(1024, 5)
      LogPrint ("response: " & sResp)
   'response_display
      If Len(sResp > 1) And Left(sResp, 2) = "qk" Then
        sResp = Mid(sResp, 2)
      End If
      If sResp = "" Or Left(sResp, 1) <> "k" Then
        LogPrint "K.O. - communication failed..."
        FlushTCP
        TryConnectTCP = 1 ' ko - retry
      Else
        If Len(sResp) > 1 Then
            LogPrint "OK!! - firmware version is " & Mid(sResp, 2)
            If Mid(sResp, 2, 3) = "KNX" Then
               gate = 2
               frmKnxScsGate.Caption = MAINCAPTION + " - KNX MODE"
            Else
               gate = 1
               frmKnxScsGate.Caption = MAINCAPTION + " - SCS MODE"
            End If
        Else
            LogPrint "OK!! "
        End If
        TryConnectTCP = 0 ' ok
      End If
    
      TryConnectTCP = 0
      Exit Function
    End If
TryError:
    MsgBox "TCP connection error - restart the app"
    TryConnectTCP = 1
End Function

Private Sub DisconnectTCP()
'    gate = 0
'    frmKnxScsGate.Caption = "KnxScsGate"
    Call Winsock1_Close
End Sub

Function WriteBufTCP(sString)
     tcpTimeout = 0
     Winsock1.SendData (sString)
End Function

Function ReadWaitTCP(length, timeout)
    Dim tim, Chread  ', strData, ln
    Dim nrbytes
    ReadWaitTCP = ""
    tim = 0
    Do
    
        DequeueTCP
        
        If (tcpBufferLen > 0) Then
            Chread = tcpBufferIn ' Leggo il contenuto del buffer di ricezione (e svuoto)
            tcpBufferIn = ""
            tcpBufferLen = 0
        Else
            Chread = ""
        End If
        If Chread = "" Then
            tim = tim + 1
        Else
            SerBuffer = SerBuffer & Chread
        End If
    Loop Until (Len(SerBuffer) >= length Or tim >= 50000 * timeout) ' 5 secondi timeout
    If Len(SerBuffer) > length Then
        ReadWaitTCP = Left(SerBuffer, length)
        SerBuffer = Mid(SerBuffer, length + 1)
    Else
        ReadWaitTCP = SerBuffer
        SerBuffer = ""
    End If
End Function

Function ReadNoWaitTCP(length)
    Dim tim, Chread ', strData, ln
    Dim nrbytes
        ReadNoWaitTCP = ""
        tim = 0

        DequeueTCP

       'Timer1_Timer
        If (tcpBufferLen > 0) Then
            Chread = tcpBufferIn ' Leggo il contenuto del buffer di ricezione (e svuoto)
            tcpBufferIn = ""
            tcpBufferLen = 0
        Else
            Chread = ""
        End If

        If Chread = "" Then
            tim = tim + 1
        Else
            SerBuffer = SerBuffer & Chread
        End If
    If Len(SerBuffer) > length Then
        ReadNoWaitTCP = Left(SerBuffer, length)
        SerBuffer = Mid(SerBuffer, length + 1)
    Else
        ReadNoWaitTCP = SerBuffer
        SerBuffer = ""
    End If
End Function

Function DequeueTCP()
' tutto viene fatto in DataArrival...
'
'
'    Dim strData, ln
' =======================================================
'        strData = ""
        DoEvents
'        If (Winsock.State = sckOpen) Or (Winsock.State = sckListening) Or (Winsock.State = sckConnected) Then
'            Winsock1.GetData strData, vbString
'            ln = Len(strData)
'            tcpBufferIn = tcpBufferIn & strData
'            tcpBufferLen = tcpBufferLen + ln
'        End If
' =======================================================
    DequeueTCP = ""
End Function

Function FlushTCP()
    Dim iCol, Chread, SerBufferHex
    Chread = tcpBufferIn ' Leggo il contenuto del buffer di ricezione (e svuoto .Input)
    tcpBufferIn = ""
    SerBuffer = SerBuffer & Chread
    If Len(SerBuffer) > 0 Then
       LogPrint Len(SerBuffer) & " bytes flushed: "
        For iCol = 1 To Len(SerBuffer)
            SerBufferHex = SerBufferHex & HexString(Asc(Mid(SerBuffer, iCol, 1)), 2)
        Next
        LogPrint (SerBufferHex)
        SerBuffer = ""
    End If
End Function

