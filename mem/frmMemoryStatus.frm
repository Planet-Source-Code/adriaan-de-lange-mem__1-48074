VERSION 5.00
Begin VB.Form frmMemoryStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Status"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmMemoryStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000002&
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   999
      Left            =   4200
      Top             =   2520
   End
End
Attribute VB_Name = "frmMemoryStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartTime
Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Initialize()
Dim lpBuffer As MemoryStatuss
GlobalMemoryStatus lpBuffer
StartTime = "Started Time : " & Format(CDbl(Now()), "dd/mm/yy hh:mm:ss") & vbNewLine

adltextstr = ""
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & StartTime
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Last Updated : " & Format(CDbl(Now()), "dd/mm/yy hh:mm:ss") & vbNewLine
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Availible Physical Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Total Physical Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Used Physical Memeory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Percentage Physical Memeory : " & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#%") & vbNewLine
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Availible Page File Size : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPageFile / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Total Page File Size : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPageFile / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Used Page File Size : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPageFile - lpBuffer.dwAvailPageFile) / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Percentage Page File Size : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPageFile / lpBuffer.dwTotalPageFile), "##.#%") & vbNewLine
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Availible Virtual Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailVirtual / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Total Virtual Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalVirtual / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Used Virtual Memeory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalVirtual - lpBuffer.dwAvailVirtual) / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Percentage Virtual Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailVirtual / lpBuffer.dwTotalVirtual), "##.#%") & vbNewLine
adltextstr = adltextstr & vbNewLine

Text1.Text = adltextstr

adlfilestr = ""
adlfilestr = adlfilestr & "!" & Format(CDbl(Now()), "dd/mm/yy hh:mm:ss") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPageFile / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwTotalPageFile / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl((lpBuffer.dwTotalPageFile - lpBuffer.dwAvailPageFile) / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPageFile / lpBuffer.dwTotalPageFile), "##.#") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailVirtual / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwTotalVirtual / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl((lpBuffer.dwTotalVirtual - lpBuffer.dwAvailVirtual) / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailVirtual / lpBuffer.dwTotalVirtual), "##.#") & "!"

Open "c:\TESTFILE.txt" For Output Shared As #1

Write #1, adlfilestr

End Sub

Private Sub Timer1_Timer()
Dim lpBuffer As MemoryStatuss
GlobalMemoryStatus lpBuffer

adltextstr = ""
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & StartTime
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Last Updated : " & Format(CDbl(Now()), "dd/mm/yy hh:mm:ss") & vbNewLine
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Availible Physical Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Total Physical Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Used Physical Memeory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Percentage Physical Memeory : " & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#%") & vbNewLine
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Availible Page File Size : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPageFile / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Total Page File Size : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPageFile / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Used Page File Size : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPageFile - lpBuffer.dwAvailPageFile) / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Percentage Page File Size : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPageFile / lpBuffer.dwTotalPageFile), "##.#%") & vbNewLine
adltextstr = adltextstr & vbNewLine

adltextstr = adltextstr & "Availible Virtual Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailVirtual / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Total Virtual Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalVirtual / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Used Virtual Memeory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalVirtual - lpBuffer.dwAvailVirtual) / 1048576), "#.## MB") & vbNewLine
adltextstr = adltextstr & "Percentage Virtual Memeory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailVirtual / lpBuffer.dwTotalVirtual), "##.#%") & vbNewLine
adltextstr = adltextstr & vbNewLine

Text1.Text = adltextstr

adlfilestr = ""
adlfilestr = adlfilestr & "!" & Format(CDbl(Now()), "dd/mm/yy hh:mm:ss") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPageFile / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwTotalPageFile / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl((lpBuffer.dwTotalPageFile - lpBuffer.dwAvailPageFile) / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailPageFile / lpBuffer.dwTotalPageFile), "##.#") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailVirtual / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwTotalVirtual / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl((lpBuffer.dwTotalVirtual - lpBuffer.dwAvailVirtual) / 1048576), "#.##") & "!"
adlfilestr = adlfilestr & Format(CDbl(lpBuffer.dwAvailVirtual / lpBuffer.dwTotalVirtual), "##.#") & "!"

Write #1, adlfilestr

End Sub
