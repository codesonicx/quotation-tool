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
Public Const GRAND_TOTALS_ROW As Long = 2
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

' Section / assembly columns
Public Const HEADER_SECTION_NAME As String = "Section Name:"
Public Const HEADER_ITEM As String = "Item"
Public Const HEADER_ITEM_DES As String = "Item Des."
Public Const HEADER_ITEM_BUILD As String = "Item Build"

' Component columns
Public Const HEADER_COMP_LIST As String = "Comp List"
Public Const HEADER_COMPONENT As String = "Component"
Public Const HEADER_COMP_DES As String = "Comp Des."
Public Const HEADER_COMP_BUILD As String = "Comp Build"
Public Const HEADER_SECT_DES As String = "Sect Des."
Public Const HEADER_SECT_BUILD As String = "Sect Build"

' Markup columns
Public Const HEADER_MATERIAL_MARKUP As String = "Material Markup [%]"
Public Const HEADER_LABOR_MARKUP As String = "Labor Markup [%]"

' Hours / rate columns
Public Const HEADER_DESIGN_HOURS As String = "Design [h]"
Public Const HEADER_CONSTRUCTION_HOURS As String = "Construction [h]"
Public Const HEADER_DEBUG_HOURS As String = "Debug/Tryout [h]"
Public Const HEADER_INSTALLATION_HOURS As String = "Installation [h]"
Public Const HEADER_STARTUP_HOURS As String = "Start up @ Client [h]"
Public Const HEADER_MATERIAL_RATE As String = "Material Rate [$/Build]"

' Budgeted cost columns
Public Const HEADER_BUDGETED_DESIGN_COST As String = "Budgeted Design Cost [$]"
Public Const HEADER_BUDGETED_CONSTRUCTION_COST As String = "Budgeted Construction Cost [$]"
Public Const HEADER_BUDGETED_DEBUG_COST As String = "Budgeted Debug/Tryout Cost [$]"
Public Const HEADER_BUDGETED_INSTALLATION_COST As String = "Budgeted Installation Cost [$]"
Public Const HEADER_BUDGETED_STARTUP_COST As String = "Budgeted Start up @ Client Cost [$]"
Public Const HEADER_BUDGETED_MATERIAL_COST As String = "Budgeted Material Cost [$]"
Public Const HEADER_BUDGETED_TOTAL As String = "Budgeted Total [$]"

' Markup cost columns
Public Const HEADER_DESIGN_COST_WITH_MARKUP As String = "Design Cost w/ Markup [$]"
Public Const HEADER_CONSTRUCTION_COST_WITH_MARKUP As String = "Construction Cost w/ Markup [$]"
Public Const HEADER_DEBUG_COST_WITH_MARKUP As String = "Debug/Tryout Cost w/ Markup [$]"
Public Const HEADER_INSTALLATION_COST_WITH_MARKUP As String = "Installation Cost w/ Markup [$]"
Public Const HEADER_STARTUP_COST_WITH_MARKUP As String = "Start up @ Client Cost w/ Markup [$]"
Public Const HEADER_MATERIAL_COST_WITH_MARKUP As String = "Material Cost w/ Markup [$]"
Public Const HEADER_TOTAL_WITH_MARKUP As String = "Total w/ Markup [$]"

' Final percentage column
Public Const HEADER_PERCENT As String = "%"

' =========================================================
' Fixed worksheet columns
' =========================================================

Public Const COL_ITEM As Long = 2
Public Const COL_ITEM_DES As Long = 3
Public Const COL_ITEM_BUILD As Long = 4

Public Const COL_COMP_LIST As Long = 5
Public Const COL_COMPONENT As Long = 6
Public Const COL_COMP_DES As Long = 7
Public Const COL_COMP_BUILD As Long = 8

Public Const COL_SECT_DES As Long = 9
Public Const COL_SECT_BUILD As Long = 10

Public Const COL_MATERIAL_MARKUP As Long = 11
Public Const COL_LABOR_MARKUP As Long = 12

Public Const COL_DESIGN_HOURS As Long = 13
Public Const COL_CONSTRUCTION_HOURS As Long = 14
Public Const COL_DEBUG_HOURS As Long = 15
Public Const COL_INSTALLATION_HOURS As Long = 16
Public Const COL_STARTUP_HOURS As Long = 17

Public Const COL_MATERIAL_RATE As Long = 18

Public Const COL_BUDGET_DESIGN As Long = 19
Public Const COL_BUDGET_CONSTRUCTION As Long = 20
Public Const COL_BUDGET_DEBUG As Long = 21
Public Const COL_BUDGET_INSTALLATION As Long = 22
Public Const COL_BUDGET_STARTUP As Long = 23
Public Const COL_BUDGET_MATERIAL As Long = 24
Public Const COL_BUDGET_TOTAL As Long = 25

Public Const COL_MARKUP_DESIGN As Long = 26
Public Const COL_MARKUP_CONSTRUCTION As Long = 27
Public Const COL_MARKUP_DEBUG As Long = 28
Public Const COL_MARKUP_INSTALLATION As Long = 29
Public Const COL_MARKUP_STARTUP As Long = 30
Public Const COL_MARKUP_MATERIAL As Long = 31
Public Const COL_MARKUP_TOTAL As Long = 32

Public Const COL_PERCENT As Long = 33

' =========================================================
' Number formats
' =========================================================
Public Const FORMAT_INTEGER As String = "#,##0"
Public Const FORMAT_DECIMAL As String = "#,##0.0"
Public Const FORMAT_CURRENCY As String = "$#,##0.00"
Public Const FORMAT_PERCENT As String = "0%"

' =========================================================
' Component list sheets
' =========================================================
Public Const SHEET_ME_COMPONENT_LIST As String = "ME Component List"
Public Const SHEET_CE_COMPONENT_LIST As String = "CE Component List"

Public Const COMPONENT_LIST_FIRST_DATA_ROW As Long = 2

Public Const COL_COMPONENT_LIST_ID As Long = 1
Public Const COL_COMPONENT_LIST_ITEM As Long = 2
Public Const COL_COMPONENT_LIST_DESIGN As Long = 3
Public Const COL_COMPONENT_LIST_CONSTRUCTION As Long = 4
Public Const COL_COMPONENT_LIST_DEBUG As Long = 5
Public Const COL_COMPONENT_LIST_INSTALLATION As Long = 6
Public Const COL_COMPONENT_LIST_STARTUP As Long = 7
Public Const COL_COMPONENT_LIST_MATERIAL_RATE As Long = 8

' =========================================================
' Component list table names
' =========================================================
Public Const TABLE_ME_COMPONENT_LIST As String = "CompList"
Public Const TABLE_CE_COMPONENT_LIST As String = "CVCompList"

' =========================================================
' Summary sheet
' =========================================================
Public Const SHEET_SUMMARY As String = "Summary"
Public Const TABLE_SUMMARY As String = "SummaryTable"
Public Const SUMMARY_HEADER_TAB_NAME As String = "Tab Name"

' =========================================================
' Fixed workbook sheets
' =========================================================
Public Const SHEET_COVER_PAGE As String = "Cover Page"
Public Const SHEET_SETUP As String = "Set-Up"
Public Const SHEET_UPDATING_SHEET As String = "Updating sheet"
Public Const SHEET_DRIVERS As String = "Drivers"
Public Const SHEET_NOTES As String = "Notes"