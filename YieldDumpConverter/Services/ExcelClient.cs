using System;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace YieldDumpConverter.Services
{
    class ExcelClient
    {
        /// <summary>
        /// Opens a new Excel window and pastes the text into the default worksheet.
        /// </summary>
        /// <param name="textToPaste">Text to be pasted into the Excel</param>
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

                // Copy and paste the textbox content to the active worksheet.
                Clipboard.Clear();
                Clipboard.SetText(textToPaste);
                xlWorkbook.ActiveSheet.Paste();

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

                    // In debug mode, the EXCEL.EXE process may still be alive. Kill it by the window handle.
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
