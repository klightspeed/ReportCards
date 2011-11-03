using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace SouthernCluster.ReportCards
{
    /// <summary>
    /// Interaction logic for WPFGUI.xaml
    /// </summary>
    public partial class WPFGUI : Window
    {
        private Thread MergerThread;

        public WPFGUI()
        {
            InitializeComponent();
        }

        public void AddRecord(string name)
        {
            CheckBox cbRecord = new CheckBox();
            cbRecord.IsChecked = true;
            cbRecord.Content = name;
            cbRecord.Checked += clbRecord_Checked;
            cbRecord.Unchecked += clbRecord_Unchecked;
            clbRecords.Items.Add(cbRecord);
        }

        public void ClearRecords()
        {
            clbRecords.Items.Clear();
        }

        private string GetNextMergeRecord()
        {
            foreach (object o in clbRecords.Items)
            {
                if (o is CheckBox)
                {
                    CheckBox cbRecord = (CheckBox)o;
                    if (cbRecord.IsChecked.HasValue && cbRecord.IsChecked.Value)
                    {
                        cbRecord.IsChecked = false;
                        return (string)cbRecord.Content;
                    }
                }
            }
            return null;
        }
        
        private void MergeRecords(bool PDF, bool PUB, bool Print)
        {
            List<string> records = new List<string>();
            foreach (object o in clbRecords.Items)
            {
                if (o is CheckBox)
                {
                    CheckBox cbRecord = (CheckBox)o;
                    if (cbRecord.IsChecked.HasValue && cbRecord.IsChecked.Value)
                    {
                        records.Add((string)cbRecord.Content);
                    }
                }
            }
        }
        
        private void btnMerge_Click(object sender, RoutedEventArgs e)
        {
            MergeRecords(
                cbPDF.IsChecked.HasValue && cbPDF.IsChecked.Value,
                cbPublisher.IsChecked.HasValue && cbPublisher.IsChecked.Value,
                false
            );
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            MergeRecords(false, false, true);
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnTemplate_Click(object sender, RoutedEventArgs e)
        {

        }

        private void tbTemplate_TextInput(object sender, TextCompositionEventArgs e)
        {

        }

        private void tbDatasource_TextInput(object sender, TextCompositionEventArgs e)
        {

        }

        private void btnDatasource_Click(object sender, RoutedEventArgs e)
        {

        }

        private void tbSaveTo_TextInput(object sender, TextCompositionEventArgs e)
        {

        }

        private void btnSaveTo_Click(object sender, RoutedEventArgs e)
        {

        }

        private void cbSelectAll_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void cbSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void clbRecord_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void clbRecord_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void tbTemplate_Drop(object sender, DragEventArgs e)
        {

        }

        private void tbTemplate_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private void tbDatasource_Drop(object sender, DragEventArgs e)
        {

        }

        private void tbDatasource_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private void tbTemplate_PreviewDragEnter(object sender, DragEventArgs e)
        {

        }
    }
}
