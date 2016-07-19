using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using YieldDumpConverter.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace YieldDumpConverter.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            textBoxMain.Focus();
            DataObject.AddPastingHandler(textBoxMain, OnPaste);
        }

        /// <summary>
        /// Intercepts the clipboard text on paste event and runs it through the
        /// conversion process.
        /// </summary>
        private void OnPaste(object sender, DataObjectPastingEventArgs e)
        {
            bool isText = e.SourceDataObject.GetDataPresent(DataFormats.UnicodeText, true);
            if (!isText)
                return;

            string text = e.SourceDataObject.GetData(DataFormats.Text) as string;
            text = Converter.Convert(text);
            DataObject d = new DataObject();
            d.SetData(DataFormats.Text, text);
            e.DataObject = d;
        }

        /// <summary>
        /// Opens a new Excel window and pastes the text into the worksheet.
        /// </summary>
        /// <param name="textToPaste">Text to be pasted into the Excel</param>
        private void OpenInExcel(string textToPaste)
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
                Clipboard.SetText(textBoxMain.Text);
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

        private void btnCrash_Click(object sender, RoutedEventArgs e)
        {
            throw new Exception();
        }

        private void OpenInExcelCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = !string.IsNullOrWhiteSpace(textBoxMain.Text);
        }

        private void OpenInExcelCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                OpenInExcel(textBoxMain.Text);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
