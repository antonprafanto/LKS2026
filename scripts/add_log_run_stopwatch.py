"""Add one-click stopwatch macros and buttons to Log Run Otonom sheet."""

from __future__ import annotations

import sys
import winreg
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_XLSX = ROOT / "templates" / "LKS2026-Penilaian-Juri.xlsx"
DEFAULT_XLSM = ROOT / "templates" / "LKS2026-Penilaian-Juri.xlsm"
SHEET = "Log Run Otonom"
MODULE_NAME = "LogRunStopwatch"

VBA_CODE = '''Option Explicit

Private stopwatchRunning As Boolean
Private stopwatchStart As Date

Public Sub CatatSIAP()
    With Worksheets("Log Run Otonom").Range("B8")
        .Value = Time
        .NumberFormat = "hh:mm:ss"
    End With
End Sub

Public Sub CatatSTART()
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    With ws.Range("E8")
        .Value = Time
        .NumberFormat = "hh:mm:ss"
    End With
    stopwatchStart = Now
    stopwatchRunning = True
    ws.Range("K7").Value = 0
    ws.Range("K7").NumberFormat = "[h]:mm:ss"
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch", , False
    On Error GoTo 0
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch"
End Sub

Public Sub CatatSELESAI()
    stopwatchRunning = False
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    With ws.Range("H8")
        .Value = Time
        .NumberFormat = "hh:mm:ss"
    End With
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch", , False
    On Error GoTo 0
End Sub

Public Sub LogRun_TickStopwatch()
    If Not stopwatchRunning Then Exit Sub
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    If ws.Range("H8").Value <> "" Then
        stopwatchRunning = False
        Exit Sub
    End If
    If ws.Range("E8").Value = "" Then Exit Sub
    ws.Range("K7").Value = Now - stopwatchStart
    ws.Range("K7").NumberFormat = "[h]:mm:ss"
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch"
End Sub

Public Sub ResetWaktuRun()
    stopwatchRunning = False
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "LogRun_TickStopwatch", , False
    On Error GoTo 0
    Dim ws As Worksheet
    Set ws = Worksheets("Log Run Otonom")
    ws.Range("B8").ClearContents
    ws.Range("E8").ClearContents
    ws.Range("H8").ClearContents
    ws.Range("K7").ClearContents
End Sub
'''

BUTTONS = [
    ("CatatSIAP", "▶ SIAP"),
    ("CatatSTART", "▶ START"),
    ("CatatSELESAI", "■ SELESAI"),
    ("ResetWaktuRun", "↺ Reset"),
]


def _enable_vba_project_access() -> None:
    """Allow Excel COM to inject VBA (build-time only on dev machine)."""
    for version in ("16.0", "15.0", "14.0"):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{version}\Excel\Security",
                0,
                winreg.KEY_SET_VALUE,
            )
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            return
        except OSError:
            continue


def _remove_existing_module(vb_project, name: str) -> None:
    for comp in vb_project.VBComponents:
        if comp.Name == name:
            vb_project.VBComponents.Remove(comp)
            return


def _remove_form_buttons(ws) -> None:
    for i in range(ws.Shapes.Count, 0, -1):
        shape = ws.Shapes(i)
        if shape.Type in (8, 12):  # xlFormControl, xlButton
            try:
                if shape.FormControlType == 0:  # xlButtonControl
                    shape.Delete()
            except Exception:
                pass


def apply_stopwatch(
    xlsx_path: Path = DEFAULT_XLSX,
    xlsm_path: Path = DEFAULT_XLSM,
    *,
    visible: bool = False,
) -> Path:
    """Build .xlsm with stopwatch buttons from base .xlsx."""
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("pywin32 required: pip install pywin32") from exc

    if not xlsx_path.exists():
        raise FileNotFoundError(xlsx_path)

    _enable_vba_project_access()

    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = visible
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(str(xlsx_path.resolve()))
        ws = wb.Worksheets(SHEET)

        _remove_form_buttons(ws)
        vb_project = wb.VBProject
        _remove_existing_module(vb_project, MODULE_NAME)
        mod = vb_project.VBComponents.Add(1)
        mod.Name = MODULE_NAME
        mod.CodeModule.AddFromString(VBA_CODE)

        anchor = ws.Range("J8")
        left = anchor.Left
        top = anchor.Top
        width = 58
        height = 22
        gap = 4
        for idx, (macro, caption) in enumerate(BUTTONS):
            btn = ws.Shapes.AddFormControl(0, left + idx * (width + gap), top, width, height)
            btn.TextFrame.Characters().Text = caption
            btn.OnAction = macro

        if xlsm_path.exists():
            xlsm_path.unlink()
        wb.SaveAs(str(xlsm_path.resolve()), FileFormat=52)
        return xlsm_path
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


def main() -> int:
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_XLSX
    xlsm = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_XLSM
    try:
        out = apply_stopwatch(xlsx, xlsm)
        print(f"Created stopwatch workbook: {out}")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        print(
            "Tip: Excel > File > Options > Trust Center > Macro Settings > "
            "Trust access to the VBA project object model",
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
