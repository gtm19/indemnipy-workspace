from dataclasses import asdict
from pathlib import Path
from typing import Any

import pytest
from indemnipy_ai.capabilities.excel._functions import (
    _oletools_vba_parser as oletools_vba_parser,
)
from indemnipy_ai.capabilities.excel._workbook_protocol import VbaSummary
from inline_snapshot import snapshot

TEST_DATA_DIR = Path(__file__).parent.parent.parent / "data"


@pytest.mark.parametrize(
    "file_path,expected",
    [
        (
            TEST_DATA_DIR / "Vantris_Pharmaceuticals_STP_Submission_2026.xlsm",
            snapshot(
                {
                    "analysis_results": [
                        {
                            "kw_type": "Suspicious",
                            "keyword": "Write",
                            "description": "May write to a file (if combined with Open)",
                        },
                        {
                            "kw_type": "Suspicious",
                            "keyword": "Run",
                            "description": "May run an executable file or a system command",
                        },
                        {
                            "kw_type": "Suspicious",
                            "keyword": "CreateObject",
                            "description": "May create an OLE object",
                        },
                        {
                            "kw_type": "Suspicious",
                            "keyword": "Hex Strings",
                            "description": "Hex-encoded strings were detected, may be used to obfuscate strings (option --decode to see all)",
                        },
                    ],
                    "macros": [
                        {
                            "filename": "xl/vbaProject.bin",
                            "stream_path": "VBA/ThisWorkbook",
                            "vba_filename": "ThisWorkbook.cls",
                            "vba_code": """\
Attribute VB_Name = "ThisWorkbook"\r
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
Option Explicit\r
\r
'====================================================\r
' 1. Calculate Total Insured Value (TIV) across locations\r
'====================================================\r
Sub CalculateTotalInsuredValue()\r
    Dim ws As Worksheet\r
    Dim lastRow As Long\r
    Dim i As Long\r
    Dim totalTIV As Double\r
    \r
    Set ws = ThisWorkbook.Sheets("Locations") ' adjust sheet name as needed\r
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\r
    totalTIV = 0\r
    \r
    For i = 2 To lastRow ' assumes row 1 = headers\r
        totalTIV = totalTIV + ws.Cells(i, "D").Value ' assumes col D = Stock Value\r
    Next i\r
    \r
    ws.Range("D" & lastRow + 2).Value = "Total Insured Value:"\r
    ws.Range("E" & lastRow + 2).Value = totalTIV\r
    ws.Range("E" & lastRow + 2).NumberFormat = "#,##0.00"\r
    \r
    MsgBox "Total Insured Value calculated: " & Format(totalTIV, "#,##0.00"), vbInformation\r
End Sub\r
\r
'====================================================\r
' 2. Flag locations exceeding a PML / per-location limit\r
'====================================================\r
Sub FlagLocationsOverLimit()\r
    Dim ws As Worksheet\r
    Dim lastRow As Long\r
    Dim i As Long\r
    Dim perLocationLimit As Double\r
    \r
    Set ws = ThisWorkbook.Sheets("Locations")\r
    perLocationLimit = InputBox("Enter the per-location limit to check against:", _\r
                                 "PML Threshold", 5000000)\r
    If perLocationLimit <= 0 Then Exit Sub\r
    \r
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\r
    \r
    For i = 2 To lastRow\r
        If ws.Cells(i, "D").Value > perLocationLimit Then\r
            ws.Cells(i, "D").Interior.Color = RGB(255, 199, 206) ' light red\r
            ws.Cells(i, "D").Font.Color = RGB(156, 0, 6)\r
        Else\r
            ws.Cells(i, "D").Interior.ColorIndex = xlNone\r
            ws.Cells(i, "D").Font.ColorIndex = xlAutomatic\r
        End If\r
    Next i\r
    \r
    MsgBox "Locations exceeding " & Format(perLocationLimit, "#,##0") & " have been highlighted.", vbInformation\r
End Sub\r
\r
'====================================================\r
' 3. Check for missing required fields before submission\r
'====================================================\r
Sub ValidateSubmissionData()\r
    Dim ws As Worksheet\r
    Dim lastRow As Long\r
    Dim i As Long\r
    Dim missingCount As Long\r
    Dim missingList As String\r
    Dim requiredCols As Variant\r
    Dim c As Variant\r
    \r
    Set ws = ThisWorkbook.Sheets("Locations")\r
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\r
    requiredCols = Array("A", "B", "C", "D") ' e.g., Location Name, Address, Product Type, Stock Value\r
    missingCount = 0\r
    missingList = ""\r
    \r
    For i = 2 To lastRow\r
        For Each c In requiredCols\r
            If Trim(ws.Cells(i, c).Value & "") = "" Then\r
                missingCount = missingCount + 1\r
                missingList = missingList & "Row " & i & ", Col " & c & vbCrLf\r
            End If\r
        Next c\r
    Next i\r
    \r
    If missingCount > 0 Then\r
        MsgBox "Missing data found in " & missingCount & " field(s):" & vbCrLf & missingList, vbExclamation\r
    Else\r
        MsgBox "All required fields are complete.", vbInformation\r
    End If\r
End Sub\r
\r
'====================================================\r
' 4. Generate a quick summary by product/temperature category\r
' (useful for cold-chain pharma STP submissions)\r
'====================================================\r
Sub SummarizeByStorageType()\r
    Dim ws As Worksheet\r
    Dim summaryWs As Worksheet\r
    Dim lastRow As Long\r
    Dim i As Long\r
    Dim dict As Object\r
    Dim key As Variant\r
    \r
    Set ws = ThisWorkbook.Sheets("Locations")\r
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row\r
    Set dict = CreateObject("Scripting.Dictionary")\r
    \r
    ' Assumes column E = Storage Type (e.g., "Ambient", "Refrigerated 2-8C", "Frozen")\r
    For i = 2 To lastRow\r
        key = ws.Cells(i, "E").Value\r
        If Not dict.exists(key) Then\r
            dict.Add key, ws.Cells(i, "D").Value\r
        Else\r
            dict(key) = dict(key) + ws.Cells(i, "D").Value\r
        End If\r
    Next i\r
    \r
    ' Write summary to a new sheet\r
    On Error Resume Next\r
    Application.DisplayAlerts = False\r
    ThisWorkbook.Sheets("Summary").Delete\r
    Application.DisplayAlerts = True\r
    On Error GoTo 0\r
    \r
    Set summaryWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))\r
    summaryWs.Name = "Summary"\r
    summaryWs.Range("A1").Value = "Storage Type"\r
    summaryWs.Range("B1").Value = "Total Value"\r
    \r
    Dim r As Long\r
    r = 2\r
    For Each key In dict.Keys\r
        summaryWs.Cells(r, "A").Value = key\r
        summaryWs.Cells(r, "B").Value = dict(key)\r
        summaryWs.Cells(r, "B").NumberFormat = "#,##0.00"\r
        r = r + 1\r
    Next key\r
    \r
    summaryWs.Columns("A:B").AutoFit\r
    MsgBox "Summary sheet created.", vbInformation\r
End Sub\r
\r
'====================================================\r
' 5. Run all checks before finalizing submission\r
'====================================================\r
Sub RunFullSubmissionCheck()\r
    ValidateSubmissionData\r
    CalculateTotalInsuredValue\r
    FlagLocationsOverLimit\r
    SummarizeByStorageType\r
End Sub\r
\r
\r
""",
                        },
                        {
                            "filename": "xl/vbaProject.bin",
                            "stream_path": "VBA/Sheet4",
                            "vba_filename": "Sheet4.cls",
                            "vba_code": """\
Attribute VB_Name = "Sheet4"\r
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
\r
""",
                        },
                        {
                            "filename": "xl/vbaProject.bin",
                            "stream_path": "VBA/Sheet3",
                            "vba_filename": "Sheet3.cls",
                            "vba_code": """\
Attribute VB_Name = "Sheet3"\r
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
\r
""",
                        },
                        {
                            "filename": "xl/vbaProject.bin",
                            "stream_path": "VBA/Sheet2",
                            "vba_filename": "Sheet2.cls",
                            "vba_code": """\
Attribute VB_Name = "Sheet2"\r
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
\r
""",
                        },
                        {
                            "filename": "xl/vbaProject.bin",
                            "stream_path": "VBA/Sheet1",
                            "vba_filename": "Sheet1.cls",
                            "vba_code": """\
Attribute VB_Name = "Sheet1"\r
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r
Attribute VB_GlobalNameSpace = False\r
Attribute VB_Creatable = False\r
Attribute VB_PredeclaredId = True\r
Attribute VB_Exposed = True\r
Attribute VB_TemplateDerived = False\r
Attribute VB_Customizable = True\r
\r
""",
                        },
                    ],
                }
            ),
        ),
        (
            TEST_DATA_DIR / "Vantris_Pharmaceuticals_STP_Submission_2026.xlsx",
            snapshot(
                {
                    "analysis_results": [],
                    "macros": [],
                }
            ),
        ),
    ],
)
def test_vba_extraction(file_path: Path, expected: dict[str, Any]):
    summary = VbaSummary.from_file(file_path, parser=oletools_vba_parser)
    d = asdict(summary)
    assert d.pop("filepath") == file_path
    assert d == expected
