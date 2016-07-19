using System;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace YieldDumpConverter.Services
{
    class ExcelClient
    {
        /// <summary>
        /// Opens a new Excel window and pastes the passed yield dump text into the default worksheet.
        /// </summary>
        /// <param name="textToPaste">Raw yield dump to be pasted into the Excel</param>
        public static void OpenInExcel(string textToPaste)
        {
            try
            {
                OpenInExcelInternal(textToPaste);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void OpenInExcelInternal(string textToPaste)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;

            try
            {
                xlApp = new Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Unable to open Excel.");
                    return;
                }

                // Create a new empty workbook.
                xlWorkbook = xlApp.Workbooks.Add();

                // Copy and paste the passed text to the active worksheet.
                // This has a side effect of changing the clipboard content.
                // Maybe there is a way to do this without using clipboard.
                Clipboard.Clear();
                Clipboard.SetText(textToPaste);
                xlWorkbook.ActiveSheet.Paste();

                // Increase column widths to fit all text nicely.
                xlWorkbook.ActiveSheet.Columns[1].ColumnWidth = 20;
                xlWorkbook.ActiveSheet.Columns[2].ColumnWidth = 15;

                // Make the Excel window visible to the user.
                xlApp.Visible = true;
            }
            catch (Exception ex)
            {
                // Release hidden references.
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false, Type.Missing, Type.Missing);
                }

                if (xlApp != null)
                {
                    int hwnd = xlApp.Application.Hwnd;
                    xlApp.Quit();

                    // In debug mode, EXCEL.EXE process may still be alive after we have released all handles.
                    // Kill it by the window handle.
                    Utility.TryKillProcessByMainWindowHwnd(hwnd);
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                // Release all handles now so that we don't leave dangling EXCEL.EXE processes
                // after the user has closed the Excel window.
                Marshal.FinalReleaseComObject(xlWorkbook);
                Marshal.FinalReleaseComObject(xlApp);
            }
        }

    }
}
