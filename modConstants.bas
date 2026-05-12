Option Explicit

' =========================================================
' Workbook / worksheet names
' =========================================================
Public Const SHEET_TEMPLATE As String = "Template"
Public Const SHEET_ASSEMBLY_LIST As String = "Assembly List"

' =========================================================
' Domains
' =========================================================
Public Const DOMAIN_ME As String = "ME"
Public Const DOMAIN_CE As String = "CE"

' =========================================================
' Library / template labels
' =========================================================
Public Const TEMPLATE_ME_MANUAL_ENTRY As String = "ME - Manual Entry"

' =========================================================
' General row / column structure
' =========================================================
Public Const FIRST_DATA_ROW As Long = 3

Public Const COL_SECTION_NAME As Long = 1        ' Column A
Public Const COL_ASSEMBLY_NAME As Long = 2       ' Column B
Public Const COL_LIBRARY_COLOR As Long = 1       ' Column A
Public Const COL_LIBRARY_NAME As Long = 2        ' Column B
Public Const COL_COPY_START As Long = 2          ' Column B

Public Const COL_SECTION_TOTAL_START As Long = 13    ' Column M
Public Const COL_SECTION_TOTAL_END As Long = 32      ' Column AF

' =========================================================
' Row colors
' =========================================================
Public Const COLOR_INDEX_SECTION_BLUE As Long = 37
Public Const COLOR_INDEX_ASSEMBLY_GRAY As Long = 15

Public Const COLOR_ASSEMBLY_GRAY_R As Long = 191
Public Const COLOR_ASSEMBLY_GRAY_G As Long = 191
Public Const COLOR_ASSEMBLY_GRAY_B As Long = 191

' =========================================================
' Header labels
' =========================================================
Public Const HEADER_COMPONENT As String = "Component"
Public Const HEADER_COMP_DES As String = "Comp Des."
Public Const HEADER_COMP_BUILD As String = "Comp Build"
Public Const HEADER_SECT_DES As String = "Sect Des."
Public Const HEADER_SECT_BUILD As String = "Sect Build"
Public Const HEADER_DESIGN_HOURS As String = "design [h]"
Public Const HEADER_CONSTRUCTION_HOURS As String = "construction [h]"
Public Const HEADER_DEBUG_HOURS As String = "debug/tryout [h]"
Public Const HEADER_INSTALLATION_HOURS As String = "installation [h]"
Public Const HEADER_STARTUP_HOURS As String = "start up @ client [h]"
Public Const HEADER_MATERIAL_RATE As String = "Material Rate [$/Build]"
Public Const HEADER_COMP_LIST As String = "Comp List"

Public Const HEADER_ITEM As String = "Item"
Public Const HEADER_ITEM_DES As String = "Item" & vbLf & "Des."
Public Const HEADER_ITEM_BUILD As String = "Item" & vbLf & "Build"
