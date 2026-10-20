"""Print the structure of an Excel workbook without printing any cell values.

Usage (Windows, Command Prompt):
    python inspect_workbook.py "C:\\path\\to\\report.xlsx"

Opens the file read-only through Excel, prints sheet names, table ranges,
header labels, formulas, PivotTables, charts and named ranges, then closes it.
Nothing is modified and nothing is sent anywhere. Review the output before
sharing it: header labels and formulas are shown as-is.
"""
import sys

import win32com.client


def main(path):
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
    try:
        print(f"Python {sys.version.split()[0]} | Excel {excel.Version}")
        print(f"File: {wb.Name}")
        print(f"Sheets: {wb.Worksheets.Count}   Charts-only sheets: {wb.Chartsheets.Count}")

        print("\n== Named ranges ==")
        for n in wb.Names:
            print(f"  {n.Name} -> {n.RefersTo}")

        for ws in wb.Worksheets:
            ur = ws.UsedRange
            print(f"\n=== Sheet '{ws.Name}' (visible={ws.Visible == -1}) ===")
            print(f"  Used range: {ur.Address(False, False)}  rows={ur.Rows.Count} cols={ur.Columns.Count}")

            for lo in ws.ListObjects:
                print(f"  Table '{lo.Name}': {lo.Range.Address(False, False)}")
                print(f"    Columns: {[c.Name for c in lo.ListColumns]}")

            print("  First 3 rows (text cells only, formulas shown as formulas):")
            for r in range(ur.Row, ur.Row + min(3, ur.Rows.Count)):
                cells = []
                for c in range(ur.Column, ur.Column + min(ur.Columns.Count, 30)):
                    cell = ws.Cells(r, c)
                    f = cell.Formula
                    if isinstance(f, str) and f.startswith("="):
                        cells.append(f)
                    elif isinstance(cell.Value, str):
                        cells.append(cell.Value)
                    else:
                        cells.append("<num/date/blank>")
                print(f"    row {r}: {cells}")

            # Formulas in the first data row below the header, to show patterns.
            if ur.Rows.Count > 1:
                forms = []
                for c in range(ur.Column, ur.Column + min(ur.Columns.Count, 30)):
                    f = ws.Cells(ur.Row + 1, c).Formula
                    if isinstance(f, str) and f.startswith("="):
                        forms.append(f"{ws.Cells(ur.Row + 1, c).Address(False, False)}: {f}")
                if forms:
                    print("  Formulas in first data row:")
                    for f in forms:
                        print(f"    {f}")

            for pt in ws.PivotTables():
                print(f"  PivotTable '{pt.Name}' at {pt.TableRange2.Address(False, False)}")
                print(f"    Source: {pt.SourceData}")
                print(f"    Rows: {[f.Name for f in pt.RowFields]}")
                print(f"    Cols: {[f.Name for f in pt.ColumnFields]}")
                print(f"    Data: {[f.Name for f in pt.DataFields]}")

            for co in ws.ChartObjects():
                ch = co.Chart
                print(f"  Chart '{co.Name}' type={ch.ChartType} at {co.TopLeftCell.Address(False, False)}")
                for s in ch.SeriesCollection():
                    print(f"    Series: {s.Formula}")

            print(f"  Pictures/shapes: {ws.Shapes.Count} total")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
