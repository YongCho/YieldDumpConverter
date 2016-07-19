using System;
using System.Windows;
using System.Windows.Input;
using YieldDumpConverter.Services;

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
            ExcelClient.OpenInExcel(textBoxMain.Text);
        }
    }
}
