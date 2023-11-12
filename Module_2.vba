Option Explicit
Private IOTagCol, IODscCol, IOTypCol, IOAddCol, IOAlmEnCol, IOHistEnCol, IODTCol, IOEUMaxCol, IOEUMinCol, IOEUMaxRCol, IOEUMinRCol, IORMaxCol, IORMinCol, _
 IOEUCol, IOAreaCol, IOAlmPriCol, IOOnCol, IOOffCol, IOAlmTypCol, IOContCol, IOContNameCol, IOPagCol, IOAttribNameCol, _
 IOAlmHiCol, IOAlmHiHiCol, IOAlmLoCol, IOAlmLoLoCol, IOPriHiCol, IOPriHiHiCol, IOPriLoCol, IOPriLoLoCol, _
 IOAlmHiEnCol, IOAlmHiHiEnCol, IOAlmLoEnCol, IOAlmLoLoEnCol, IOAlmInhibitCol, IOCmdDataCol, IORWCol, IOAlmDBCol, IODlyCol As Integer
Private PLC, ScanGroup As String
Private Extension(27) As String
Private c As Object
Private d As Object
Private e As String
Private f As Object
Private HeaderCreated As Boolean
Private ContainType As Boolean
Private GORow As Long
Private IORow As Long
Private FileName As String
Dim Attrib() As String
Dim FieldList() As String
Dim FieldListColumn() As Integer
Dim RCount As Integer
Dim ObjCnt As Integer
Dim RR(100, 100) As String
Dim GOArray(100, 20, 21) As String

Sub ShowCreateGObject()

    CreateGObject_UF.Show
    
End Sub

Sub CreateCSV()

Dim FirstRow As Integer
Dim Lastrow As Integer

Sheets("GalaxyLoad").Select
FileName = Cells(1, 3).Value
FirstRow = Cells(2, 3).Value
Lastrow = Cells(3, 3).Value
Sheets("GalaxyLoad").Rows(FirstRow & ":" & Lastrow).Select
Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:= _
    FileName _
    , FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
End Sub

Sub CreateGalaxyObjects()
Dim c As Object
Dim matchFoundIndex As Long
Dim AttrCnt As Integer

'Check for duplicate Tagname
Sheets("WWSP").Select

For Each c In Sheets("WWSP").Range("A:A")
    If c.Value <> "" Then
        matchFoundIndex = WorksheetFunction.Match(Cells(c.Row, 1), Range("A:A"), 0)
        If matchFoundIndex <> c.Row Then
            MsgBox "Duplicate Tagname: " & Cells(matchFoundIndex, 1).Value
            MsgBox (matchFoundIndex)
            Exit Sub
            Exit For
        End If
   End If
Next

'Check list if empty
Sheets("WWPivot").Select
RCount = Sheets("WWPivot").Cells(1, 2) + 9

If RCount = 9 Then
    MsgBox "Instrument Tagname column is empty."
    Exit Sub
End If

'Clear sheet
Sheets("GalaxyLoad").Cells.Clear
Sheets("GalaxyLoad").Rows("1:7").Delete

'Copy Header
Sheets("GalaxyTemplates").Select
Sheets("GalaxyTemplates").Rows("1:7").Select
Selection.Copy
Sheets("GalaxyLoad").Select
Cells(1, 1).Select
ActiveSheet.Paste

'Get WWSP Columns Number
Call Get_IO_Columns

PLC = Sheets("IO List").Cells(2, 6)
ScanGroup = Sheets("IO List").Cells(2, 7)


Sheets("GalaxyLoad").Cells(2, 3).Value = Cells(Rows.Count, "A").End(xlUp).Row + 2


Call TemplateToObjects("$s00_Bad_Data_Alm_01")
Call TemplateToObjects("$s00_CB_00")
Call TemplateToObjects("$m_IO_FLOAT")
Call TemplateToObjects("$s00_IOBOOL_00")
Call TemplateToObjects("$s00_IOBOOL_01")
Call TemplateToObjects("$s00_IODINT_01")
Call TemplateToObjects("$s00_IOFLOAT_01")
Call TemplateToObjects("$s00_IOINT_00")
Call TemplateToObjects("$s00_IOINT_01")
Call TemplateToObjects("$s00_IOSTRING_00")
Call TemplateToObjects("$s00_PLC_Comm_00")
Call TemplateToObjects("$s00_RFAB_AB_Loop_01")
Call TemplateToObjects("$s01_RFAB_AB_Loop_00")
Call TemplateToObjects("$s00_RFAB_AB_Pump_00")
Call TemplateToObjects("$s00_RFAB_AB_VFD_00")
Call TemplateToObjects("$s00_RFAB_AB_VVD_00")
Call TemplateToObjects("$s00_RFAB_AB_VVD_01")
Call TemplateToObjects("$s00_RFAB_AI_Limits_00")
Call TemplateToObjects("$s00_RFAB_AI_Limits_01")
Call TemplateToObjects("$s00_RFAB_AI_Limits_03")
Call TemplateToObjects("$s00_RFAB_AI_Limits_07")
Call TemplateToObjects("$s00_RFAB_AI_Limits_08")
Call TemplateToObjects("$s00_RFAB_AI_Limits_09")
Call TemplateToObjects("$s00_RFAB_AI_LSS_00")
Call TemplateToObjects("$s01_RFAB_AI_LSS_00")
Call TemplateToObjects("$s00_RFAB_ALM_DIS_00")
Call TemplateToObjects("$s00_RFAB_AMC_00")
Call TemplateToObjects("$s00_RFAB_AMC_AI_00")
Call TemplateToObjects("$s00_RFAB_CB_02")
Call TemplateToObjects("$s00_RFAB_CDU_Tank_00")
Call TemplateToObjects("$s00_RFAB_DOSING_PMP_00")
Call TemplateToObjects("$s00_RFAB_DOSING_PMP_01")
Call TemplateToObjects("$s00_RFAB_Drum_00")
Call TemplateToObjects("$s00_RFAB_FFU_00")
Call TemplateToObjects("$s00_RFAB_Float_Control_00")
Call TemplateToObjects("$s00_RFAB_Gas_Cabinet_00")
Call TemplateToObjects("$s00_RFAB_Leak_Detect_00")
Call TemplateToObjects("$s00_RFAB_LIEBERT_DS_00")
Call TemplateToObjects("$s00_RFAB_Maint_00")
Call TemplateToObjects("$s00_RFAB_PID_00")
Call TemplateToObjects("$s01_RFAB_PID_00")
Call TemplateToObjects("$s00_RFAB_PLC_Diagnostics_LSS_00")
Call TemplateToObjects("$s01_RFAB_Pump_00")
Call TemplateToObjects("$s03_RFAB_Pump_00")
Call TemplateToObjects("$s03_RFAB_Pump_01")
Call TemplateToObjects("$S04_RFAB_Pump_00")
Call TemplateToObjects("$s05_RFAB_Pump_00")
Call TemplateToObjects("$s00_RFAB_QC_SC_MEAS_01")
Call TemplateToObjects("$s00_RFAB_QC_SC_MEAS_02")
Call TemplateToObjects("$s00_RFAB_Reset_00")
Call TemplateToObjects("$s00_RFAB_S7_AUTOJMPR_01")
Call TemplateToObjects("$s00_RFAB_S7_AvgCalc_01")
Call TemplateToObjects("$s01_RFAB_S7_BYPASS_01")
Call TemplateToObjects("$s00_RFAB_S7_BYPASS_01")
Call TemplateToObjects("$s00_RFAB_S7_CB_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_CB_Mode_001")
Call TemplateToObjects("$s00_RFAB_S7_CB_Mode_01")
Call TemplateToObjects("$s00_RFAB_S7_CDIW_LOOP_MODE_00")
Call TemplateToObjects("$s00_RFAB_S7_CH_40STS_01")
Call TemplateToObjects("$s00_RFAB_S7_CH_FLUSH_01")
Call TemplateToObjects("$s00_RFAB_S7_CH_FTR_01")
Call TemplateToObjects("$s00_RFAB_S7_CH_STS_00")
Call TemplateToObjects("$s00_RFAB_S7_CHEM_MODE_00")
Call TemplateToObjects("$s00_RFAB_S7_CMP_LOOP_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_DIG_MON_01")
Call TemplateToObjects("$s00_RFAB_S7_DIG_MON_02")
Call TemplateToObjects("$s00_RFAB_S7_EMO_01")
Call TemplateToObjects("$s00_RFAB_S7_GPUMP_00")
Call TemplateToObjects("$s00_RFAB_S7_GPUMP_01")
Call TemplateToObjects("$s00_RFAB_S7_HDIW_LOOP_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_InSelect_01")
Call TemplateToObjects("$s00_RFAB_S7_IWWR_DISTRIB_00")
Call TemplateToObjects("$s00_RFAB_S7_LITETREE_01")
Call TemplateToObjects("$s01_RFAB_S7_LITETREE_00")
Call TemplateToObjects("$s01_RFAB_S7_LITETREE_01")
Call TemplateToObjects("$s00_RFAB_S7_LS_01")
Call TemplateToObjects("$s01_RFAB_S7_MIDASAT_01")
Call TemplateToObjects("$s00_RFAB_S7_MOD_VLV_01")
Call TemplateToObjects("$s02_RFAB_S7_MOD_VLV_00")
Call TemplateToObjects("$s00_RFAB_S7_OP_A_01")
Call TemplateToObjects("$s00_RFAB_S7_OP_A_LIM_03")
Call TemplateToObjects("$s00_RFAB_S7_OP_D_01")
Call TemplateToObjects("$s00_RFAB_S7_OP_I_LIM_00")
Call TemplateToObjects("$m_RFAB_S7_OP_TRIG_00")
Call TemplateToObjects("$m_RFAB_S7_OP_TRIG_01")
Call TemplateToObjects("$s00_RFAB_S7_PMB_MODE_00")
Call TemplateToObjects("$s00_RFAB_S7_PMB_Mux_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_POMB_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_PS_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_RIWW_CB_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_RIWW_RO1_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_RO_Mode_01")
Call TemplateToObjects("$s00_RFAB_S7_RO1_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_RO2_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_ROReject_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_S120VFD_00")
Call TemplateToObjects("$s00_RFAB_S7_S120VFD_FC_00")
Call TemplateToObjects("$$s00_RFAB_S7_SC_AGGR_00")
Call TemplateToObjects("$s00_RFAB_S7_SC_AGGR_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_DIGM8_ST_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_LMN3P_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_MEAS_01")
Call TemplateToObjects("$s00_S7_SC_MEAS_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_MOT1_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_OP_D_00")
Call TemplateToObjects("$s00_RFAB_S7_SC_OP_D_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_POLY_00")
Call TemplateToObjects("$s00_RFAB_S7_SCACCU_00")
Call TemplateToObjects("$s00_RFAB_S7_SCACCU_01")
Call TemplateToObjects("$s00_RFAB_S7_SCACCU_02")
Call TemplateToObjects("$s00_RFAB_S7_SCMEAS_SW_00")
Call TemplateToObjects("$s00_RFAB_S7_SCMEAS_SW_01")
Call TemplateToObjects("$s00_RFAB_S7_SCMEAS_SW_02")
Call TemplateToObjects("$s00_RFAB_S7_SDALARM_01")
Call TemplateToObjects("$s00_RFAB_S7_SEPIX_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_SIMODIR_00")
Call TemplateToObjects("$s00_RFAB_S7_TI_GL_FL_00")
Call TemplateToObjects("$s00_RFAB_S7_TI_GL_FL_01")
Call TemplateToObjects("$s00_RFAB_S7_TI_PID_01")
Call TemplateToObjects("$s00_RFAB_S7_TI_PID_02")
Call TemplateToObjects("$s00_RFAB_S7_VSD_02")
Call TemplateToObjects("$s02_RFAB_S7_VSD_03")
Call TemplateToObjects("$s02_RFAB_S7_VSD_04")
Call TemplateToObjects("$s02_RFAB_S7_VSD_05")
Call TemplateToObjects("$s00_RFAB_S7_VSD_FPOS_00")
Call TemplateToObjects("$s00_RFAB_S7_VSD_FPOS_01")
Call TemplateToObjects("$s00_RFAB_S7_WARN_SDI_01")
Call TemplateToObjects("$s00_RFAB_S7_WARN_SDI_02")
Call TemplateToObjects("$m_RFAB_S7_WireBrk_00")
Call TemplateToObjects("$s00_RFAB_SD_LSS_00")
Call TemplateToObjects("$s00_RFAB_SD_LSS_01")
Call TemplateToObjects("$s00_RFAB_SOV_00")
Call TemplateToObjects("$s01_RFAB_SOV_00")
Call TemplateToObjects("$s02_RFAB_SOV_00")
Call TemplateToObjects("$s00_RFAB_Tank_00")
Call TemplateToObjects("$s00_RFAB_Tank_02")
Call TemplateToObjects("$s00_RFAB_Tank_04")
Call TemplateToObjects("$m_RFAB_TI_PID_TIA_00")
Call TemplateToObjects("$s00_RFAB_TI_PID_TIA_01")
Call TemplateToObjects("$s00_RFAB_VFD_00")
Call TemplateToObjects("$s01_RFAB_VFD_00")
Call TemplateToObjects("$s00_RFAB_VMB_STICK_00")
Call TemplateToObjects("$s01_RFAB_VMB_STICK_00")
Call TemplateToObjects("$s00_RFAB_VMB_STICK_02")
Call TemplateToObjects("$s00_RFAB_VVD_00")
Call TemplateToObjects("$s00_RFAB_S7_OP_TRIG_01")
Call TemplateToObjects("$s00_S7_CPU_Diag_00")
Call TemplateToObjects("$s00_S7_IO_Diag_00")
Call TemplateToObjects("$s00_SFC_00")
Call TemplateToObjects("$s00_Stale_Data_Alm_00")
Call TemplateToObjects("$s00_RFAB_AI_Limits_07")
Call TemplateToObjects("$s00_RFAB_CB_00")
Call TemplateToObjects("$s00_RFAB_CB_01")
Call TemplateToObjects("$s00_RFAB_Isotainer_00")
Call TemplateToObjects("$s00_RFAB_MODE_00")
Call TemplateToObjects("$s00_RFAB_Pump_00")
Call TemplateToObjects("$s00_RFAB_Reset_00")
Call TemplateToObjects("$s00_RFAB_S7_InSelect_01")
Call TemplateToObjects("$s00_RFAB_S7_MB_Mode_00")
Call TemplateToObjects("$s00_RFAB_S7_MIDASAT_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_AGGR_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_LMN3P_01")
Call TemplateToObjects("$s00_RFAB_S7_SC_OP_D_01")
Call TemplateToObjects("$s00_RFAB_S7_SCACCU_00")
Call TemplateToObjects("$s00_RFAB_S7_SMC_DIR_01")
Call TemplateToObjects("$s00_RFAB_S7_SMC_MEAS_00")
Call TemplateToObjects("$s00_RFAB_S7_VDD_01")
Call TemplateToObjects("$s00_RFAB_Tank_01")
Call TemplateToObjects("$s00_RFAB_TIME_00")
Call TemplateToObjects("$s00_RFAB_Vaporizer_00")
Call TemplateToObjects("$s00_RFAB_VMB_MODE_00")
Call TemplateToObjects("$s00_RFAB_XFMR_00")
Call TemplateToObjects("$s00_Wago_hmiValve_00")
Call TemplateToObjects("$s01_RFAB_PID_00")
Call TemplateToObjects("$s01_RFAB_PLC_Status_00")
Call TemplateToObjects("$s01_RFAB_S7_FT_00")
Call TemplateToObjects("$s02_Wago_VVD_00")
Call TemplateToObjects("$s00_RFAB_S7_OP_A_LIM_02")

'====Create Remmote Response Objects====
ReDim FieldList(8, 1) As String
ReDim FieldListColumn(8, 1) As Integer
Call Create_RemoteResponse

If CreateGObject_UF.PLC_Object.Value = True Then
    Call Create_DIObjects
End If

If CreateGObject_UF.Area_Object.Value = True Then
    Call Create_AreaObjects
End If
End Sub
Sub Get_IO_Columns()

Dim d As Object

'Get WWSP Columns Number
For Each d In Sheets("WWSP").Range("9:9")
    If d.Value = "Object Name" Then
        IOTagCol = d.Column
    ElseIf d.Value = "Description" Then
        IODscCol = d.Column
    ElseIf d.Value = "Template" Then
        IOTypCol = d.Column
    ElseIf d.Value = "Input Source" Then
        IOAddCol = d.Column
    ElseIf d.Value = "Attribute Type" Then
        IODTCol = d.Column
    ElseIf d.Value = "EngUnits Min" Then
        IOEUMinCol = d.Column
    ElseIf d.Value = "EngUnits Max" Then
        IOEUMaxCol = d.Column
    ElseIf d.Value = "EngUnits Range Min" Then
        IOEUMinRCol = d.Column
    ElseIf d.Value = "EngUnits Range Max" Then
        IOEUMaxRCol = d.Column
    ElseIf d.Value = "Raw Min" Then
        IORMinCol = d.Column
    ElseIf d.Value = "Raw Max" Then
        IORMaxCol = d.Column
    ElseIf d.Value = "Eng Units" Then
        IOEUCol = d.Column
    ElseIf d.Value = "Area" Then
        IOAreaCol = d.Column
    ElseIf d.Value = "Alarm Ext" Then
        IOAlmEnCol = d.Column
    ElseIf d.Value = "Alm Priority" Then
        IOAlmPriCol = d.Column
    ElseIf d.Value = "Dig Stat 0" Then
        IOOffCol = d.Column
    ElseIf d.Value = "Dig Stat 1" Then
        IOOnCol = d.Column
    ElseIf d.Value = "Alm Type" Then
        IOAlmTypCol = d.Column
    ElseIf d.Value = "Container" Then
        IOContCol = d.Column
    ElseIf d.Value = "Contained Name" Then
        IOContNameCol = d.Column
    ElseIf d.Value = "Paging Enable" Then
        IOPagCol = d.Column
    ElseIf d.Value = "Hi Limit" Then
        IOAlmHiCol = d.Column
    ElseIf d.Value = "HiHi Limit" Then
        IOAlmHiHiCol = d.Column
    ElseIf d.Value = "Lo Limit" Then
        IOAlmLoCol = d.Column
    ElseIf d.Value = "LoLo Limit" Then
        IOAlmLoLoCol = d.Column
    ElseIf d.Value = "Hi Priority" Then
        IOPriHiCol = d.Column
    ElseIf d.Value = "HiHi Priority" Then
        IOPriHiHiCol = d.Column
    ElseIf d.Value = "Lo Priority" Then
        IOPriLoCol = d.Column
    ElseIf d.Value = "LoLo Priority" Then
        IOPriLoLoCol = d.Column
    ElseIf d.Value = "Hi Alarmed" Then
        IOAlmHiEnCol = d.Column
    ElseIf d.Value = "HiHi Alarmed" Then
        IOAlmHiHiEnCol = d.Column
    ElseIf d.Value = "Lo Alarmed" Then
        IOAlmLoEnCol = d.Column
    ElseIf d.Value = "LoLo Alarmed" Then
        IOAlmLoLoEnCol = d.Column
    ElseIf d.Value = "Hist Ext" Then
        IOHistEnCol = d.Column
    ElseIf d.Value = "Attribute Name" Then
        IOAttribNameCol = d.Column
    ElseIf d.Value = "Alarm Inhibit" Then
        IOAlmInhibitCol = d.Column
    ElseIf d.Value = "Read/ Write" Then
        IORWCol = d.Column
    ElseIf d.Value = "Deadband" Then
        IOAlmDBCol = d.Column
    ElseIf d.Value = "Delay" Then
        IODlyCol = d.Column
    End If
   
Next

End Sub

Sub Create_AreaObjects()
Dim AreaCount As Integer
Dim Lastrow As Long
Dim TArea1Row As Long
Dim AreaRow As Long

'Check list if empty.
Sheets("Area List").Select
AreaCount = Sheets("Area List").Cells(Rows.Count, "B").End(xlUp).Row - 3

If AreaCount <= 0 Then
    MsgBox "Area Name column is empty."
    Exit Sub
End If
 
'Delete old area objects and find last row
Sheets("GalaxyLoad").Select
For Each c In Sheets("GalaxyLoad").Range("A:A")
    If c = ":TEMPLATE=$s00_Location" Then
        Lastrow = c.Row
        Sheets("GalaxyLoad").Rows(c.Row).Select
        Range(Selection, Rows(c.Row).End(xlDown)).Delete Shift:=xlUp
    End If
Next
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row + 2

'Find Area template row.
Sheets("GalaxyTemplates").Select
For Each c In Sheets("GalaxyTemplates").Range("A:A")
    If c = ":TEMPLATE=$Location" Then
        TArea1Row = c.Row
    End If
Next

'Create Galaxy Object from AREA template
HeaderCreated = False
For AreaRow = 4 To AreaCount + 3
        If HeaderCreated = False Then
            'Create Headers
            HeaderCreated = True
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(TArea1Row & ":" & TArea1Row + 1).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(Lastrow, 1).Select
            ActiveSheet.Paste
            Lastrow = Lastrow + 1
        End If
        Lastrow = Lastrow + 1
        Sheets("GalaxyTemplates").Select
        Sheets("GalaxyTemplates").Rows(TArea1Row + 2).Select
        Selection.Copy
        Sheets("GalaxyLoad").Select
        Cells(Lastrow, 1).Select
        ActiveSheet.Paste
        'Area
        Cells(Lastrow, 1).Value = Sheets("Area List").Cells(AreaRow, 2)
        Cells(Lastrow, 6).Value = Sheets("Area List").Cells(AreaRow, 3)
        Cells(Lastrow, 2).Value = Sheets("Area List").Cells(AreaRow, 4)
        Cells(Lastrow, 4).Value = Sheets("Area List").Cells(AreaRow, 5)
Next
            
Sheets("GalaxyLoad").Cells(3, 3).Value = Sheets("GalaxyLoad").Cells(Rows.Count, "A").End(xlUp).Row
FileName = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "."))
Sheets("GalaxyLoad").Cells(1, 3).Value = FileName & "csv"
Sheets("GalaxyLoad").Select

End Sub

Sub Create_DIObjects()
Dim PLCCount As Integer
Dim Lastrow As Long
Dim T1Row As Long
Dim T2Row As Long
Dim T3Row As Long
Dim T4Row As Long
Dim T5Row As Long
Dim T6Row As Long
Dim PLCRow As Long
Dim Row2 As Long
Dim Row3 As Long

'Check list if empty.
Sheets("IO List").Select

If Cells(2, 6) = "" Then
    MsgBox "PLC Name is empty."
    Exit Sub
End If

Sheets("GalaxyLoad").Select
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row + 2

'Find PLC template row.
Sheets("GalaxyTemplates").Select
For Each c In Sheets("GalaxyTemplates").Range("A:A")
    If c = "ControlLogixCIP" Then
        T1Row = c.Row + 1
    ElseIf c = "WagoPLC" Then
        T2Row = c.Row + 1
    ElseIf c = "S7_1200" Then
        T3Row = c.Row + 1
    ElseIf c = "S7_300" Then
        T4Row = c.Row + 1
    ElseIf c = "S7_400" Then
        T5Row = c.Row + 1
    End If
Next

'Create PLC Objects from Template
'Tempalate 1 : ControlLogix CIP
'Ethernet Module
If Sheets("IO List").Cells(2, 9).Value = "ControlLogixCIP" Then
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(T1Row & ":" & T1Row + 10).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(Lastrow, 1).Select
            ActiveSheet.Paste
            Cells(Lastrow + 2, 1).Value = Sheets("IO List").Cells(2, 6) & "_ENB"
            Cells(Lastrow + 2, 4).Value = Sheets("IO List").Cells(2, 10)
            Cells(Lastrow + 2, 5).Value = Sheets("IO List").Cells(2, 8) & " ethernet bridge module"
            Cells(Lastrow + 2, 6).Value = Sheets("IO List").Cells(2, 11)
            Cells(Lastrow + 6, 1).Value = Sheets("IO List").Cells(2, 6) & "_BP"
            Cells(Lastrow + 6, 4).Value = Sheets("IO List").Cells(2, 6) & "_ENB"
            Cells(Lastrow + 6, 5).Value = Sheets("IO List").Cells(2, 8) & " backplane"
            Cells(Lastrow + 10, 1).Value = Sheets("IO List").Cells(2, 6)
            Cells(Lastrow + 10, 4).Value = Sheets("IO List").Cells(2, 6) & "_BP"
            Cells(Lastrow + 10, 5).Value = Sheets("IO List").Cells(2, 8)
End If
        
If Sheets("IO List").Cells(2, 9).Value = "WagoPLC" Then
'Template 2 : Wago
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(T2Row & ":" & T2Row + 2).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(Lastrow, 1).Select
            ActiveSheet.Paste
            Cells(Lastrow + 2, 1).Value = Sheets("IO List").Cells(2, 6)
            Cells(Lastrow + 2, 4).Value = Sheets("IO List").Cells(2, 10)
            Cells(Lastrow + 2, 5).Value = Sheets("IO List").Cells(2, 8)
            Cells(Lastrow + 2, 19).Value = Sheets("IO List").Cells(2, 11)
End If
            
If Sheets("IO List").Cells(2, 9).Value = "S7_1200" Then
'Template 3 : S7-1200
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(T3Row & ":" & T3Row + 2).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(Lastrow, 1).Select
            ActiveSheet.Paste
            Cells(Lastrow + 2, 1).Value = Sheets("IO List").Cells(2, 6)
            Cells(Lastrow + 2, 4).Value = Sheets("IO List").Cells(2, 10)
            Cells(Lastrow + 2, 5).Value = Sheets("IO List").Cells(2, 8)
            Cells(Lastrow + 2, 6).Value = Sheets("IO List").Cells(2, 11)
End If
            
If Sheets("IO List").Cells(2, 9).Value = "S7_300" Then
'Template 4 : S7-300
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(T4Row & ":" & T4Row + 2).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(Lastrow, 1).Select
            ActiveSheet.Paste
            Cells(Lastrow + 2, 1).Value = Sheets("IO List").Cells(2, 6)
            Cells(Lastrow + 2, 4).Value = Sheets("IO List").Cells(2, 10)
            Cells(Lastrow + 2, 5).Value = Sheets("IO List").Cells(2, 8)
            Cells(Lastrow + 2, 6).Value = Sheets("IO List").Cells(2, 11)
End If
            
If Sheets("IO List").Cells(2, 9).Value = "S7_400" Then
'Template 5 : S7-400
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(T5Row & ":" & T5Row + 2).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(Lastrow, 1).Select
            ActiveSheet.Paste
            Cells(Lastrow + 2, 1).Value = Sheets("IO List").Cells(2, 6)
            Cells(Lastrow + 2, 4).Value = Sheets("IO List").Cells(2, 10)
            Cells(Lastrow + 2, 5).Value = Sheets("IO List").Cells(2, 8)
            Cells(Lastrow + 2, 6).Value = Sheets("IO List").Cells(2, 11)
End If

Sheets("GalaxyLoad").Cells(3, 3).Value = Sheets("GalaxyLoad").Cells(Rows.Count, "A").End(xlUp).Row
FileName = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "."))
Sheets("GalaxyLoad").Cells(1, 3).Value = FileName & "csv"
Sheets("GalaxyLoad").Select

End Sub
Sub Create_RemoteResponse()
Dim ColumnFound As Boolean
Dim TemplateMarker As String
Dim TRow As Long
Dim AlmEn As String
Dim AlmPri As String
Dim AlmTyp As String
Dim y As Integer
Dim z As String
Dim RRCount As Integer
Dim LastCol As Integer
Dim ObjName As String
Dim LmtAlm As Integer

'Proceed only if "WWSP" contains this type of instrument
Sheets("WWSP").Select
Sheets("WWSP").Columns(IOPagCol).Select
If Selection.Find("Y") Is Nothing Then
    Sheets("GalaxyLoad").Select
    Exit Sub
End If

'Get Template Row Number
TemplateMarker = ":TEMPLATE=$s00_RemoteNotifier"
For Each c In Sheets("GalaxyTemplates").Range("A:A")
    If c = TemplateMarker Then
        TRow = c.Row
    End If
Next

'Get Template Columns Number
Sheets("GalaxyTemplates").Select
Sheets("GalaxyTemplates").Rows(TRow + 1).Select
LastCol = Sheets("GalaxyTemplates").Cells(TRow + 1, Columns.Count).End(xlToLeft).Column

For Each d In Selection
    ColumnFound = False
    If d.Value = "" Then
        Exit For
    ElseIf d.Value = "Area" Then
        FieldListColumn(1, 1) = d.Column
    ElseIf d.Value = "ShortDesc" Then
        FieldListColumn(2, 1) = d.Column
    ElseIf d.Value = "Container" Then
        FieldListColumn(3, 1) = d.Column
    ElseIf d.Value = "ContainedName" Then
        FieldListColumn(4, 1) = d.Column
    ElseIf d.Value = "TimeDelay" Then
        FieldListColumn(5, 1) = d.Column
    ElseIf d.Value = "CommunicatorObjName" Then
        FieldListColumn(6, 1) = d.Column
    ElseIf d.Value = "StateAlarm.InputSource" Then
        FieldListColumn(7, 1) = d.Column
    ElseIf d.Value = "StateAlarm.MsgFormat" Then
        FieldListColumn(8, 1) = d.Column
    End If
Next

'Find First Galaxy Object Row
Sheets("GalaxyLoad").Select
GORow = Cells(Rows.Count, "A").End(xlUp).Row + 2


'Create RR Objects from template
HeaderCreated = False
Sheets("WWSP").Select
For IORow = 10 To RCount
    If Sheets("WWSP").Cells(IORow, IOTagCol).Value <> "" Then
        ObjName = Sheets("WWSP").Cells(IORow, IOTagCol).Value
    End If
    
    z = Sheets("WWSP").Cells(IORow, IODTCol).Value
    e = Sheets("WWSP").Cells(IORow, IOPagCol).Value
    
    For y = 1 To 5
        If y = 1 And z = "BOOL" Then
            AlmPri = Sheets("WWSP").Cells(IORow, IOAlmPriCol).Value
            AlmEn = Sheets("WWSP").Cells(IORow, IOAlmEnCol).Value
            AlmTyp = "State"
        ElseIf y = 2 And (z = "REAL" Or z = "FLOAT" Or z = "INT" Or z = "WORD" Or z = "DINT" Or z = "DWORD") Then
            AlmEn = Sheets("WWSP").Cells(IORow, IOAlmHiEnCol).Value
            AlmPri = Sheets("WWSP").Cells(IORow, IOPriHiCol).Value
            AlmTyp = "H"
        ElseIf y = 3 And (z = "REAL" Or z = "FLOAT" Or z = "INT" Or z = "WORD" Or z = "DINT" Or z = "DWORD") Then
            AlmEn = Sheets("WWSP").Cells(IORow, IOAlmHiHiEnCol).Value
            AlmPri = Sheets("WWSP").Cells(IORow, IOPriHiHiCol).Value
            AlmTyp = "HH"
        ElseIf y = 4 And (z = "REAL" Or z = "FLOAT" Or z = "INT" Or z = "WORD" Or z = "DINT" Or z = "DWORD") Then
            AlmEn = Sheets("WWSP").Cells(IORow, IOAlmLoEnCol).Value
            AlmPri = Sheets("WWSP").Cells(IORow, IOPriLoCol).Value
            AlmTyp = "L"
        ElseIf y = 5 And (z = "REAL" Or z = "FLOAT" Or z = "INT" Or z = "WORD" Or z = "DINT" Or z = "DWORD") Then
            AlmEn = Sheets("WWSP").Cells(IORow, IOAlmLoLoEnCol).Value
            AlmPri = Sheets("WWSP").Cells(IORow, IOPriLoLoCol).Value
            AlmTyp = "LL"
        Else
            AlmEn = "N"
        End If
    
        If e = "Y" And (AlmEn = "Y" Or AlmEn = "TRUE") Then
            If (IORow > 1 And HeaderCreated = False) Then
                'Create Headers
                HeaderCreated = True
                Sheets("GalaxyTemplates").Select
                Sheets("GalaxyTemplates").Rows(TRow & ":" & TRow + 1).Select
                Selection.Copy
                Sheets("GalaxyLoad").Select
                Cells(GORow, 1).Select
                ActiveSheet.Paste
                GORow = GORow + 1
            End If
            
            GORow = GORow + 1
            Sheets("GalaxyTemplates").Select
            Sheets("GalaxyTemplates").Rows(TRow + 2).Select
            Selection.Copy
            Sheets("GalaxyLoad").Select
            Cells(GORow, 1).Select
            ActiveSheet.Paste
            
            'Tagname = Instrument Tagname
            Cells(GORow, 1).Value = "RR_" & ObjName & "_" & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & "_" & AlmTyp
            If AlmTyp = "State" Then
                Cells(GORow, 1).Value = "RR_" & ObjName & Sheets("WWSP").Cells(IORow, IOAttribNameCol)
            Else
                Cells(GORow, 1).Value = "RR_" & ObjName & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & "_" & AlmTyp
            End If
            
            'Area = Area
            If FieldListColumn(1, 1) > 0 Then
                Cells(GORow, FieldListColumn(1, 1)).Value = "RemoteResponse"
            End If
            
            'ShortDesc = Description
            If FieldListColumn(2, 1) > 0 Then
                Cells(GORow, FieldListColumn(2, 1)).Value = Sheets("WWSP").Cells(IORow, IODscCol) & " Remote Notifier"
            End If
            
            'Time Delay
            'If FieldListColumn(5, 1) > 0 Then
            '    If AlmPri \ 10 < 10 Then
            '        Cells(GORow, FieldListColumn(5, 1)).Value = "0000 0" & AlmPri \ 10 & ":00:00.0000"
            '    ElseIf AlmPri \ 10 < 24 Then
            '        Cells(GORow, FieldListColumn(5, 1)).Value = "0000 " & AlmPri \ 10 & ":00:00.0000"
            '    ElseIf AlmPri \ 10 < 100 Then
            '       Cells(GORow, FieldListColumn(5, 1)).Value = "0001 00:00:00.0000"
            '    End If
            'End If
            
            'Remote Response Server Object
            If FieldListColumn(6, 1) > 0 Then
               Cells(GORow, FieldListColumn(6, 1)).Value = Sheets("IO List").Cells(2, 12)
            End If
            
            'Alarm Input Source
            If FieldListColumn(7, 1) > 0 Then
                If AlmTyp = "State" Then
                    Cells(GORow, FieldListColumn(7, 1)).Value = ObjName & "." & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & ".InAlarm"
                Else
                    Cells(GORow, FieldListColumn(7, 1)).Value = ObjName & "." & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & "." & AlmTyp & ".InAlarm"
                End If
            End If
                
            'Alarm Message
            If FieldListColumn(8, 1) > 0 Then
                If AlmTyp = "State" Then
                    Cells(GORow, FieldListColumn(8, 1)).Value = "ALARM! " & Sheets("WWSP").Cells(IORow, IODscCol) & " ([N].[A]) has occurred at [U] on [T]. Priority [P]. Value is [Z:" _
                    & ObjName & "." & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & "]."
                Else
                    Cells(GORow, FieldListColumn(8, 1)).Value = "ALARM! " & Sheets("WWSP").Cells(IORow, IODscCol) & " ([N].[A]) has occurred at [U] on [T]. Priority [P]. Value is [Z:" _
                    & ObjName & "." & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & "] [Z:" & ObjName & "." & Sheets("WWSP").Cells(IORow, IOAttribNameCol) & ".EngUnits]."
                End If
            End If
        End If
    Next
Next

Sheets("GalaxyLoad").Cells(3, 3).Value = Sheets("GalaxyLoad").Cells(Rows.Count, "A").End(xlUp).Row
FileName = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "."))
Sheets("GalaxyLoad").Cells(1, 3).Value = FileName & "csv"
Sheets("GalaxyLoad").Select
    
    
End Sub

Sub WriteExtension(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)
    
If Sheets("WWSP").Cells(IRow + Offset, IODTCol) = "BOOL" Then
    Call WriteExtensionBool(Attrib, GRow, IRow, Offset, AttribIdx)
ElseIf Sheets("WWSP").Cells(IRow + Offset, IODTCol) = "STRING" Then
    Call WriteExtensionString(Attrib, GRow, IRow, Offset, AttribIdx)
Else
    Call WriteExtensionAnalog(Attrib, GRow, IRow, Offset, AttribIdx)
End If

End Sub
Sub WriteExtensionBool(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'Description = Description
If FieldListColumn(AttribIdx, 1) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 1)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'OnMsg = Dig Stat 1
If FieldListColumn(AttribIdx, 2) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOOnCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 2)).Value = Sheets("WWSP").Cells(IRow + Offset, IOOnCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 2)).Value = "True"
    End If
End If

'OffMsg = Dig Stat 0
If FieldListColumn(AttribIdx, 3) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOOffCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 3)).Value = Sheets("WWSP").Cells(IRow + Offset, IOOffCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 3)).Value = "False"
    End If
End If

'EngUnits = EGU
If FieldListColumn(AttribIdx, 15) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 15)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUCol)
End If

'History Extension
If Sheets("WWSP").Cells(IRow + Offset, IOHistEnCol) = "Y" Then
    Call WriteExtensionBoolHist(Attrib, GRow, IRow, Offset, AttribIdx)
End If

'Alarm Extension
If Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
    Call WriteExtensionBoolAlm(Attrib, GRow, IRow, Offset, AttribIdx)
End If


'IO Extension
If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> "---" Then
    Call WriteExtensionBoolIO(Attrib, GRow, IRow, Offset, AttribIdx)
End If

End Sub
Sub WriteExtensionBoolIO(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'InputSource = Internal Address
If FieldListColumn(AttribIdx, 22) > 0 Then
    Sheets("GalaxyLoad").Cells(GRow, FieldListColumn(AttribIdx, 22)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAddCol)
End If

'DiffOutputDest
If FieldListColumn(AttribIdx, 23) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 23)).Value = "False"
End If

'InvertValue
If FieldListColumn(AttribIdx, 24) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 24)).Value = "False"
End If

'Deadband
If FieldListColumn(AttribIdx, 25) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 25)).Value = 0
End If

'OutputDest
If FieldListColumn(AttribIdx, 26) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 26)).Value = "---"
End If

'Extension + IO Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> PLC & "." & ScanGroup & "." Then
        If Sheets("WWSP").Cells(IRow + Offset, IORWCol) = "Y" Then
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
            
            Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & "/>"
        Else
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
        End If
    End If
End If

End Sub
Sub WriteExtensionBoolHist(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'ValueDeadBand
If FieldListColumn(AttribIdx, 11) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 11)).Value = 0
End If

'ForceStoragePeriod
If FieldListColumn(AttribIdx, 12) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 12)).Value = 3600000
End If

'TrendHi = Hi EGU
If FieldListColumn(AttribIdx, 13) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 13)).Value = 10
End If

'TrendLo = Lo EGU
If FieldListColumn(AttribIdx, 14) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 14)).Value = 0
End If

'Hist.DescAttrName
If FieldListColumn(AttribIdx, 16) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 16)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'InterpolationType
If FieldListColumn(AttribIdx, 17) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 17)).Value = "SystemDefault"
End If

'RolloverValue
If FieldListColumn(AttribIdx, 18) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 18)).Value = 0
End If

'SampleCount
If FieldListColumn(AttribIdx, 19) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 19)).Value = 0
End If

'EnableSwingingDoor
If FieldListColumn(AttribIdx, 20) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 20)).Value = "False"
End If

'RateDeadBand
If FieldListColumn(AttribIdx, 21) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 21)).Value = 0
End If

'Extension + Hist Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOHistEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "historyextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "historyextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
    End If
End If

End Sub
Sub WriteExtensionBoolAlm(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'HasStatistics
If FieldListColumn(AttribIdx, 4) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 4)).Value = "False"
End If

'Priority = Alarm Priority
If FieldListColumn(AttribIdx, 5) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmPriCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 5)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmPriCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 5)).Value = 999
    End If
End If

'Category
If FieldListColumn(AttribIdx, 6) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 6)).Value = "Discrete"
End If

'DescAttrName
If FieldListColumn(AttribIdx, 7) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 7)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'ActiveAlarmState= Alm Typ
If FieldListColumn(AttribIdx, 8) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmTypCol) = 1 Then
        Cells(GRow, FieldListColumn(AttribIdx, 8)).Value = "TRUE"
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 8)).Value = "FALSE"
    End If
End If

'AlarmShelveCmd
If FieldListColumn(AttribIdx, 9) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 9)).Value = "Duration=0; Reason=" & Chr(34) & Chr(34) & ";"
End If

'Alarm.TimeDeadband
If FieldListColumn(AttribIdx, 10) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 10)).Value = "0000 00:00:00.0000000"
End If

'AlarmInhibit.Description
If FieldListColumn(AttribIdx, 27) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 27)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol) & " alarm inhibit"
End If

'AlarmInhibit.OnMsg
If FieldListColumn(AttribIdx, 28) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 28)).Value = "True"
End If

'AlarmInhibit.OffMsg
If FieldListColumn(AttribIdx, 29) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 29)).Value = "False"
End If

'AlarmInhibit.InputSource = Alm Inhibit
If FieldListColumn(AttribIdx, 30) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol) <> "" And Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & ".AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
        
        Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".AlarmInhibit" & Chr(34) & "/>"
        
        Cells(GRow, FieldListColumn(AttribIdx, 30)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 30)).Value = "---"
    End If
End If

'AlarmInhibit.DiffOutputDest
If FieldListColumn(AttribIdx, 31) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 31)).Value = "False"
End If

'AlarmInhibit.InvertValue
If FieldListColumn(AttribIdx, 32) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 32)).Value = "False"
End If

'OutputDest
If FieldListColumn(AttribIdx, 33) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 33)).Value = "---"
End If

'Extension + Alm Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension>"
        End If
            
        Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "booleanextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>" & _
        "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "alarmextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
       
    End If
End If
End Sub
Sub WriteExtensionAnalog(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'Description = Description
If FieldListColumn(AttribIdx, 1) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 1)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'EngUnits = EGU
If FieldListColumn(AttribIdx, 15) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 15)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUCol)
End If

'History Extension
If Sheets("WWSP").Cells(IRow + Offset, IOHistEnCol) = "Y" Then
    Call WriteExtensionAnalogHist(Attrib, GRow, IRow, Offset, AttribIdx)
End If

'Alarm Extension
If Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
    Call WriteExtensionAnalogAlm(Attrib, GRow, IRow, Offset, AttribIdx)
End If

'IO Extension
If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> "---" Then
    Call WriteExtensionAnalogIO(Attrib, GRow, IRow, Offset, AttribIdx)
End If

End Sub
Sub WriteExtensionAnalogIO(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'InputSource = Internal Address
If FieldListColumn(AttribIdx, 22) > 0 Then
    Sheets("GalaxyLoad").Cells(GRow, FieldListColumn(AttribIdx, 22)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAddCol)
End If

'DiffOutputDest
If FieldListColumn(AttribIdx, 23) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 23)).Value = "False"
End If

'Deadband
If FieldListColumn(AttribIdx, 25) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 25)).Value = 0
End If

'OutputDest
If FieldListColumn(AttribIdx, 26) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 26)).Value = "---"
End If

'===========================================================================================================
'ClampEnabled
If FieldListColumn(AttribIdx, 71) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 71)).Value = "False"
End If

'ConversionMode
If FieldListColumn(AttribIdx, 72) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 72)).Value = "Linear"
End If

'EngUnitsMax = EngUnits Max
If FieldListColumn(AttribIdx, 73) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 73)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol)
    End If
End If

'EngUnitsMin = EngUnits Min
If FieldListColumn(AttribIdx, 74) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 74)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol)
    End If
End If

'EngUnitsRangeMax = EngUnits Range Max
If FieldListColumn(AttribIdx, 75) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 75)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMaxRCol)
    End If
End If

'EngUnitsRangeMin = EngUnits Range Min
If FieldListColumn(AttribIdx, 76) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMinRCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 76)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMinRCol)
    End If
End If

'RawMax = Raw Max
If FieldListColumn(AttribIdx, 77) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IORMaxCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 77)).Value = Sheets("WWSP").Cells(IRow + Offset, IORMaxCol)
    End If
End If

'RawMax = Raw Min
If FieldListColumn(AttribIdx, 78) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IORMinCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 78)).Value = Sheets("WWSP").Cells(IRow + Offset, IORMinCol)
    End If
End If

'Extension + IO Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> PLC & "." & ScanGroup & "." Then
        If Sheets("WWSP").Cells(IRow + Offset, IORWCol) = "Y" Then
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
        Else
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
        End If
    End If
End If

'Extension + Scaling Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol) <> "" And Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol) <> "" Then
        If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> PLC & "." & ScanGroup & "." Then
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "scalingextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "scalingextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
        End If
    End If
End If

End Sub
Sub WriteExtensionAnalogAlm(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'HasStatistics
If FieldListColumn(AttribIdx, 4) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 4)).Value = "False"
End If

'LevelAlarmed
If FieldListColumn(AttribIdx, 34) > 0 Then
        If Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
            Cells(GORow, FieldListColumn(AttribIdx, 34)).Value = "True"
        Else
            Cells(GORow, FieldListColumn(AttribIdx, 34)).Value = "False"
        End If
End If

'ROCAlarmed
If FieldListColumn(AttribIdx, 35) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 35)).Value = "False"
End If

'DeviationAlarmed
If FieldListColumn(AttribIdx, 36) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 36)).Value = "False"
End If

'Hi.Alarmed = Alarm Hi Enabled
If FieldListColumn(AttribIdx, 37) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmHiEnCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 37)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmHiEnCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 37)).Value = "FALSE"
    End If
End If

'HiHi.Alarmed = Alarm HiHi Enabled
If FieldListColumn(AttribIdx, 42) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmHiHiEnCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 42)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmHiHiEnCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 42)).Value = "FALSE"
    End If
End If

'Lo.Alarmed = Alarm Lo Enabled
If FieldListColumn(AttribIdx, 43) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmLoEnCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 43)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmLoEnCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 43)).Value = "FALSE"
    End If
End If

'LoLo.Alarmed = Alarm LoLo Enabled
If FieldListColumn(AttribIdx, 44) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmLoLoEnCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 44)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmLoLoEnCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 44)).Value = "FALSE"
    End If
End If

'HiHi.Limit = Alarm Limit HiHi
If FieldListColumn(AttribIdx, 39) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmHiHiCol) <> "" Then
            Cells(GRow, FieldListColumn(AttribIdx, 39)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmHiHiCol)
    Else
        If Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol) <> "" Then
            Cells(GRow, FieldListColumn(AttribIdx, 39)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol) - 0.0001
        Else
            Cells(GRow, FieldListColumn(AttribIdx, 39)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmHiCol) + 0.0001
        End If
    End If
End If

'Hi.Limit = Alarm Limit Hi
If FieldListColumn(AttribIdx, 38) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmHiCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 38)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmHiCol)
    Else
        If Sheets("WWSP").Cells(IRow + Offset, IOAlmHiHiCol) <> "" Then
            Cells(GRow, FieldListColumn(AttribIdx, 38)).Value = Cells(GRow, FieldListColumn(AttribIdx, 39)).Value - 0.0001
        Else
            Cells(GRow, FieldListColumn(AttribIdx, 38)).Value = Cells(GRow, FieldListColumn(AttribIdx, 39)).Value - 0.0001
        End If
    End If
End If

'LoLo.Limit = Alarm Limit LoLo
If FieldListColumn(AttribIdx, 41) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmLoLoCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 41)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmLoLoCol)
    Else
        If Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol) <> "" Then
            Cells(GRow, FieldListColumn(AttribIdx, 41)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol) + 0.0001
        Else
            Cells(GRow, FieldListColumn(AttribIdx, 41)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmLoCol) - 0.0001
        End If
    End If
End If

'Lo.Limit = Alarm Limit Lo
If FieldListColumn(AttribIdx, 40) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmLoCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 40)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmLoCol)
    Else
        If Sheets("WWSP").Cells(IRow + Offset, IOAlmLoLoCol) <> "" Then
            Cells(GRow, FieldListColumn(AttribIdx, 40)).Value = Cells(GRow, FieldListColumn(AttribIdx, 41)).Value + 0.0001
        Else
            Cells(GRow, FieldListColumn(AttribIdx, 40)).Value = Cells(GRow, FieldListColumn(AttribIdx, 41)).Value + 0.0001
        End If
    End If
End If

'LevelAlarm.TimeDeadband
If FieldListColumn(AttribIdx, 45) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IODlyCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 45)).Value = Sheets("WWSP").Cells(IRow + Offset, IODlyCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 45)).Value = "00:00:00"
        End If
End If

'LevelAlarm.ValueDeadBand
If FieldListColumn(AttribIdx, 46) > 0 Then
         If Sheets("WWSP").Cells(IRow + Offset, IOAlmDBCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 46)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmDBCol)
        End If
End If

'HiHi.Priority = Alarm Priority HiHi
If FieldListColumn(AttribIdx, 47) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOPriHiHiCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 47)).Value = Sheets("WWSP").Cells(IRow + Offset, IOPriHiHiCol)
    End If
End If

'Hi.Priority = Alarm Priority Hi
If FieldListColumn(AttribIdx, 53) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOPriHiCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 53)).Value = Sheets("WWSP").Cells(IRow + Offset, IOPriHiCol)
    End If
End If

'Lo.Priority = Alarm Priority Lo
If FieldListColumn(AttribIdx, 59) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOPriLoCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 59)).Value = Sheets("WWSP").Cells(IRow + Offset, IOPriLoCol)
    End If
End If

'LoLo.Priority = Alarm Priority LoLo
If FieldListColumn(AttribIdx, 65) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOPriLoLoCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 65)).Value = Sheets("WWSP").Cells(IRow + Offset, IOPriLoLoCol)
    End If
End If

'HiHi.DescAttrName
If FieldListColumn(AttribIdx, 48) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 48)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'Hi.DescAttrName
If FieldListColumn(AttribIdx, 54) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 54)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'Lo.DescAttrName
If FieldListColumn(AttribIdx, 60) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 60)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'LoLo.DescAttrName
If FieldListColumn(AttribIdx, 66) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 66)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'HiHi.AlarmShelveCmd
If FieldListColumn(AttribIdx, 49) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 49)).Value = "Duration=0; Reason=" & Chr(34) & Chr(34) & ";"
End If

'Hi.AlarmShelveCmd
If FieldListColumn(AttribIdx, 55) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 55)).Value = "Duration=0; Reason=" & Chr(34) & Chr(34) & ";"
End If

'Lo.AlarmShelveCmd
If FieldListColumn(AttribIdx, 61) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 61)).Value = "Duration=0; Reason=" & Chr(34) & Chr(34) & ";"
End If

'LoLo.AlarmShelveCmd
If FieldListColumn(AttribIdx, 67) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 67)).Value = "Duration=0; Reason=" & Chr(34) & Chr(34) & ";"
End If

'HiHi.AlarmInhibit.Description
If FieldListColumn(AttribIdx, 50) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 50)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol) & " alarm inhibit"
End If

'Hi.AlarmInhibit.Description
If FieldListColumn(AttribIdx, 56) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 56)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol) & " alarm inhibit"
End If

'Lo.AlarmInhibit.Description
If FieldListColumn(AttribIdx, 62) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 62)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol) & " alarm inhibit"
End If

'LoLo.AlarmInhibit.Description
If FieldListColumn(AttribIdx, 68) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 68)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol) & " alarm inhibit"
End If

'HiHi.AlarmInhibit.OnMsg
If FieldListColumn(AttribIdx, 51) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 51)).Value = "True"
End If

'HiHi.AlarmInhibit.OffMsg
If FieldListColumn(AttribIdx, 52) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 52)).Value = "False"
End If

'Hi.AlarmInhibit.OnMsg
If FieldListColumn(AttribIdx, 57) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 57)).Value = "True"
End If

'Hi.AlarmInhibit.OffMsg
If FieldListColumn(AttribIdx, 58) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 58)).Value = "False"
End If

'Lo.AlarmInhibit.OnMsg
If FieldListColumn(AttribIdx, 63) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 63)).Value = "True"
End If

'Lo.AlarmInhibit.OffMsg
If FieldListColumn(AttribIdx, 64) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 64)).Value = "False"
End If

'LoLo.AlarmInhibit.OnMsg
If FieldListColumn(AttribIdx, 69) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 69)).Value = "True"
End If

'LoLo.AlarmInhibit.OffMsg
If FieldListColumn(AttribIdx, 70) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 70)).Value = "False"
End If

'Hi Alarm Inhibit
If FieldListColumn(AttribIdx, 79) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol) <> "" And Sheets("WWSP").Cells(IRow + Offset, IOAlmHiEnCol) <> "FALSE" And Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & ".Hi.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".Hi.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
        
        Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".Hi.AlarmInhibit" & Chr(34) & "/>"

        Cells(GRow, FieldListColumn(AttribIdx, 79)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol)
    End If
End If

'Hi Hi Alarm Inhibit
If FieldListColumn(AttribIdx, 83) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol) <> "" And Sheets("WWSP").Cells(IRow + Offset, IOAlmHiHiEnCol) <> "FALSE" And Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & ".HiHi.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".HiHi.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
        
        Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".HiHi.AlarmInhibit" & Chr(34) & "/>"
        
        Cells(GRow, FieldListColumn(AttribIdx, 83)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol)
    End If
End If

'Lo Alarm Inhibit
If FieldListColumn(AttribIdx, 87) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol) <> "" And Sheets("WWSP").Cells(IRow + Offset, IOAlmLoEnCol) <> "FALSE" And Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & ".Lo.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".Lo.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
        
        Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".Lo.AlarmInhibit" & Chr(34) & "/>"
        
        Cells(GRow, FieldListColumn(AttribIdx, 87)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol)
    End If
End If

'Lo Lo Alarm Inhibit
If FieldListColumn(AttribIdx, 91) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol) <> "" And Sheets("WWSP").Cells(IRow + Offset, IOAlmLoLoEnCol) <> "FALSE" And Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & ".LoLo.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".LoLo.AlarmInhibit" & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
        
        Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & ".LoLo.AlarmInhibit" & Chr(34) & "/>"
        
        Cells(GRow, FieldListColumn(AttribIdx, 91)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAlmInhibitCol)
    End If
End If

'Hi.AlarmInhibit.DiffOutputDest
If FieldListColumn(AttribIdx, 80) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 80)).Value = "False"
End If

'HiHi.AlarmInhibit.DiffOutputDest
If FieldListColumn(AttribIdx, 84) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 84)).Value = "False"
End If

'Lo.AlarmInhibit.DiffOutputDest
If FieldListColumn(AttribIdx, 88) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 88)).Value = "False"
End If

'LoLo.AlarmInhibit.DiffOutputDest
If FieldListColumn(AttribIdx, 92) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 92)).Value = "False"
End If

'Hi.AlarmInhibit.InvertValue
If FieldListColumn(AttribIdx, 81) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 81)).Value = "False"
End If

'HiHi.AlarmInhibit.InvertValue
If FieldListColumn(AttribIdx, 85) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 85)).Value = "False"
End If

'Lo.AlarmInhibit.InvertValue
If FieldListColumn(AttribIdx, 89) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 89)).Value = "False"
End If

'LoLo.AlarmInhibit.InvertValue
If FieldListColumn(AttribIdx, 93) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 93)).Value = "False"
End If

'Hi.AlarmInhibit.OutputDest
If FieldListColumn(AttribIdx, 82) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 82)).Value = "---"
End If

'HiHi.AlarmInhibit.OutputDest
If FieldListColumn(AttribIdx, 86) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 86)).Value = "---"
End If

'Lo.AlarmInhibit.OutputDest
If FieldListColumn(AttribIdx, 90) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 90)).Value = "---"
End If

'LoLo.AlarmInhibit.OutputDest
If FieldListColumn(AttribIdx, 94) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 94)).Value = "---"
End If

'Extension + Alm Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAlmEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension>"
        End If
        Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "analogextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
    End If
End If

End Sub
Sub WriteExtensionAnalogHist(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)
'ValueDeadBand
If FieldListColumn(AttribIdx, 11) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 11)).Value = 0
End If

'ForceStoragePeriod
If FieldListColumn(AttribIdx, 12) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 12)).Value = 3600000
End If

'TrendHi = Hi EGU
If FieldListColumn(AttribIdx, 13) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 13)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMaxCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 13)).Value = 100
    End If
End If

'TrendLo = Lo EGU
If FieldListColumn(AttribIdx, 14) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol) <> "" Then
        Cells(GRow, FieldListColumn(AttribIdx, 14)).Value = Sheets("WWSP").Cells(IRow + Offset, IOEUMinCol)
    Else
        Cells(GRow, FieldListColumn(AttribIdx, 14)).Value = 0
    End If
End If

'Hist.DescAttrName
If FieldListColumn(AttribIdx, 16) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 16)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'InterpolationType
If FieldListColumn(AttribIdx, 17) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 17)).Value = "SystemDefault"
End If

'RolloverValue
If FieldListColumn(AttribIdx, 18) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 18)).Value = 0
End If

'SampleCount
If FieldListColumn(AttribIdx, 19) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 19)).Value = 0
End If

'EnableSwingingDoor
If FieldListColumn(AttribIdx, 20) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 20)).Value = "False"
End If

'RateDeadBand
If FieldListColumn(AttribIdx, 21) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 21)).Value = 0
End If

'Extension + Hist Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOHistEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "historyextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "historyextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
    End If
End If

End Sub
Sub WriteExtensionString(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'Description = Description
If FieldListColumn(AttribIdx, 1) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 1)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'History Extension
If Sheets("WWSP").Cells(IRow + Offset, IOHistEnCol) = "Y" Then
    Call WriteExtensionStringHist(Attrib, GRow, IRow, Offset, AttribIdx)
End If

'IO Extension
If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> "---" Then
    Call WriteExtensionStringIO(Attrib, GRow, IRow, Offset, AttribIdx)
End If

End Sub
Sub WriteExtensionStringIO(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'InputSource = Internal Address
If FieldListColumn(AttribIdx, 22) > 0 Then
    Sheets("GalaxyLoad").Cells(GRow, FieldListColumn(AttribIdx, 22)).Value = Sheets("WWSP").Cells(IRow + Offset, IOAddCol)
End If

'DiffOutputDest
If FieldListColumn(AttribIdx, 23) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 23)).Value = "False"
End If

'OutputDest
If FieldListColumn(AttribIdx, 26) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 26)).Value = "---"
End If

'Extension + IO Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOAddCol) <> PLC & "." & ScanGroup & "." Then
        If Sheets("WWSP").Cells(IRow + Offset, IORWCol) = "Y" Then
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputoutputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
            
            Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & "/>"
        Else
            If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            Else
                Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "inputextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
            End If
        End If
    End If
End If

End Sub
Sub WriteExtensionStringHist(Attrib As String, GRow As Long, IRow As Long, Offset As Integer, AttribIdx As Integer)

'ValueDeadBand
If FieldListColumn(AttribIdx, 11) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 11)).Value = 0
End If

'ForceStoragePeriod
If FieldListColumn(AttribIdx, 12) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 12)).Value = 3600000
End If

'TrendHi = Hi EGU
If FieldListColumn(AttribIdx, 13) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 13)).Value = 10
End If

'TrendLo = Lo EGU
If FieldListColumn(AttribIdx, 14) > 0 Then
    Cells(GRow, FieldListColumn(AttribIdx, 14)).Value = 0
End If

'Hist.DescAttrName
If FieldListColumn(AttribIdx, 16) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 16)).Value = Sheets("WWSP").Cells(IRow + Offset, IODscCol)
End If

'InterpolationType
If FieldListColumn(AttribIdx, 17) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 17)).Value = "SystemDefault"
End If

'RolloverValue
If FieldListColumn(AttribIdx, 18) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 18)).Value = 0
End If

'SampleCount
If FieldListColumn(AttribIdx, 19) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 19)).Value = 0
End If

'EnableSwingingDoor
If FieldListColumn(AttribIdx, 20) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 20)).Value = "False"
End If

'RateDeadBand
If FieldListColumn(AttribIdx, 21) > 0 Then
        Cells(GRow, FieldListColumn(AttribIdx, 21)).Value = 0
End If

'Extension + Hist Ext
If FieldListColumn(6, 1) > 0 Then
    If Sheets("WWSP").Cells(IRow + Offset, IOHistEnCol) = "Y" Then
        If Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>" Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<AttributeExtension><Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "historyextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        Else
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "<Attribute Name=" & Chr(34) & Attrib & Chr(34) & " ExtensionType=" & Chr(34) & "historyextension" & Chr(34) & " InheritedFromTagName=" & Chr(34) & Chr(34) & "/>"
        End If
    End If
End If

End Sub
Sub TemplateToObjects(WWTemplateName As String)
Dim TRow, TCol, GOCol, AttribCnt, z, y, w, v As Integer
Dim c As Object
Dim matchFoundIndex As Long
Dim AttrCnt As Integer

'Proceed only if "WWSP" contains this type of instrument
If Sheets("WWSP").Range("B:B").Find(What:=WWTemplateName, LookIn:=xlValues) Is Nothing Then
    Exit Sub
End If

'Get Template Row & Column
Sheets("GalaxyTemplates").Select
For Each c In Sheets("GalaxyTemplates").Range("B:B")
    If c = WWTemplateName Then
        TRow = c.Row
        TCol = c.Column
        Exit For
    End If
Next

AttribCnt = Cells(TRow, TCol).End(xlDown).Row - TRow

ReDim FieldListColumn(AttribCnt + 7, 111) As Integer

Sheets("GalaxyLoad").Select
GORow = Cells(Rows.Count, "A").End(xlUp).Row + 2
GOCol = 1

Cells(GORow, 1).Value = ":Template=" & WWTemplateName
GORow = GORow + 1

Cells(GORow, GOCol).Value = ":Tagname"

Cells(GORow, GOCol + 1).Value = "Area"
FieldListColumn(1, 1) = 2

Cells(GORow, GOCol + 2).Value = "SecurityGroup"
FieldListColumn(2, 1) = 3

Cells(GORow, GOCol + 3).Value = "Container"
FieldListColumn(3, 1) = 4

Cells(GORow, GOCol + 4).Value = "ContainedName"
FieldListColumn(4, 1) = 5

Cells(GORow, GOCol + 5).Value = "ShortDesc"
FieldListColumn(5, 1) = 6

Cells(GORow, GOCol + 6).Value = "Extensions"
FieldListColumn(6, 1) = 7

Cells(GORow, GOCol + 7).Value = "CmdData"
FieldListColumn(7, 1) = 8

GOCol = GOCol + 8

For z = 1 To AttribCnt

    If Sheets("GalaxyTemplates").Cells(TRow + z, TCol - 1) = "BOOL" Then
        Cells(GORow, GOCol).Value = Sheets("GalaxyTemplates").Cells(TRow + z, TCol)
        FieldListColumn(z + 7, 1) = GOCol
        GOCol = GOCol + 1
        For y = 1 To 33
            If Sheets("LookupData").Cells(y + 16, 26) <> "NA" Then
                Cells(GORow, GOCol).Value = (Sheets("GalaxyTemplates").Cells(TRow + z, TCol)) & "." & Sheets("LookupData").Cells(y + 16, 26)
                FieldListColumn(z + 7, y) = GOCol
                GOCol = GOCol + 1
            End If
        Next y
    ElseIf Sheets("GalaxyTemplates").Cells(TRow + z, TCol - 1) = "STRING" Then
        Cells(GORow, GOCol).Value = Sheets("GalaxyTemplates").Cells(TRow + z, TCol)
        FieldListColumn(z + 7, 1) = GOCol
        GOCol = GOCol + 1
        For y = 1 To 26
            If Sheets("LookupData").Cells(y + 16, 28) <> "NA" Then
                Cells(GORow, GOCol).Value = (Sheets("GalaxyTemplates").Cells(TRow + z, TCol)) & "." & Sheets("LookupData").Cells(y + 16, 28)
                FieldListColumn(z + 7, y) = GOCol
                GOCol = GOCol + 1
            End If
        Next y
    Else
        Cells(GORow, GOCol).Value = Sheets("GalaxyTemplates").Cells(TRow + z, TCol)
        FieldListColumn(z + 7, 1) = GOCol
        GOCol = GOCol + 1
        For y = 1 To 94
            If Sheets("LookupData").Cells(y + 16, 27) <> "NA" Then
                Cells(GORow, GOCol).Value = Sheets("GalaxyTemplates").Cells(TRow + z, TCol) & "." & Sheets("LookupData").Cells(y + 16, 27)
                FieldListColumn(z + 7, y) = GOCol
                GOCol = GOCol + 1
            End If
        Next y
    End If
    
Next

Sheets("WWSP").Select

For IORow = 10 To RCount
    e = Sheets("WWSP").Cells(IORow, IOTypCol).Value
    If e = WWTemplateName Then
        
        GORow = GORow + 1
        
        Sheets("GalaxyLoad").Select
        Cells(GORow, 1).Select
        
        'Tagname = Instrument Tagname
        Cells(GORow, 1).Value = Sheets("WWSP").Cells(IORow, IOTagCol)
        'Area = Area
        If FieldListColumn(1, 1) > 0 Then
            Cells(GORow, FieldListColumn(1, 1)).Value = Sheets("WWSP").Cells(IORow, IOAreaCol)
        End If
        'ShortDesc = Description
        If FieldListColumn(5, 1) > 0 Then
            Cells(GORow, FieldListColumn(5, 1)).Value = Sheets("WWSP").Cells(IORow, IODscCol)
        End If
        'Container = Container
        If FieldListColumn(3, 1) > 0 Then
            Cells(GORow, FieldListColumn(3, 1)).Value = Sheets("WWSP").Cells(IORow, IOContCol)
        End If
        'Contained Name = Contained Name
        If FieldListColumn(4, 1) > 0 Then
            Cells(GORow, FieldListColumn(4, 1)).Value = Sheets("WWSP").Cells(IORow, IOContNameCol)
        End If
        
        'Extension (START)
        If FieldListColumn(6, 1) > 0 Then
            Cells(GORow, FieldListColumn(6, 1)).Value = "<ExtensionInfo><ObjectExtension/>"
        End If
        
        'CmdData (START)
        If FieldListColumn(7, 1) > 0 Then
            Cells(GORow, FieldListColumn(7, 1)).Value = "<CmdData><BooleanLabel>"
        End If
        
        'Write definition from "WWSP" to "GalaxObjects" attributes
        For w = 1 To AttribCnt
            If AttribCnt <= 1 Then
                v = 0
            Else
                v = w
            End If
            Call WriteExtension(Sheets("GalaxyTemplates").Cells(TRow + w, TCol), GORow, IORow, v, w + 7)
        Next w
        
        'Extension (END)
        If FieldListColumn(6, 1) > 0 Then
            Cells(GORow, FieldListColumn(6, 1)).Value = Cells(GORow, FieldListColumn(6, 1)).Value & "</AttributeExtension></ExtensionInfo>"
        End If
        
        'CmdData (END)
        If FieldListColumn(7, 1) > 0 Then
            If Cells(GORow, FieldListColumn(7, 1)).Value = "<CmdData><BooleanLabel>" Then
                Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "</BooleanLabel></CmdData>"
            Else
                Cells(GORow, FieldListColumn(7, 1)).Value = Cells(GORow, FieldListColumn(7, 1)).Value & "</BooleanLabel></CmdData> \n"
            End If
        End If
        
    End If
Next IORow

Sheets("GalaxyLoad").Cells(3, 3).Value = Sheets("GalaxyLoad").Cells(Rows.Count, "A").End(xlUp).Row
FileName = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "."))
Sheets("GalaxyLoad").Cells(1, 3).Value = FileName & "csv"
Sheets("GalaxyLoad").Select

End Sub

