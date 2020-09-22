Attribute VB_Name = "mdlMemoryStatus"
Public Type MemoryStatuss
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MemoryStatuss)

