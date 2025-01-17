using Kompas6API5;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportKompas
{
    public partial class LostParts : Form
    {
        private bool cancelContextMenu = false;

        public LostParts()
        {
            InitializeComponent();
        }

        private void dataGridView2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var hitTestInfo = dataGridView2.HitTest(e.X, e.Y);
                if (hitTestInfo.RowIndex >= 0 && hitTestInfo.ColumnIndex >= 0)
                {
                    dataGridView2.ClearSelection();
                    dataGridView2.Rows[hitTestInfo.RowIndex].Selected = true;
                    cancelContextMenu = false;
                }
                else
                {
                    cancelContextMenu = true;
                }
            }
        }

        private void ToolStripMenuItemOpenInKompas_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection selectedRows = dataGridView2.SelectedRows;
            foreach (DataGridViewRow selectedRow in selectedRows)
            {
                int rowIndex = selectedRow.Index;

                if (rowIndex < 0)
                {
                    continue;
                }
                ObjectAssemblyKompas objectAssemblyKompas = ReportKompas.objectsAssemblyKompas[rowIndex];
                IDocuments document = ReportKompas.application.Documents;
                document.Open(objectAssemblyKompas.FullName, true, false);
            }
        }

        private void ToolStripMenuItemOpenInExplorer_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection selectedRows = dataGridView2.SelectedRows;

            string path = ""; 

            foreach (DataGridViewRow selectedRow in selectedRows)
            {
                int rowIndex = selectedRow.Index;

                if (rowIndex < 0)
                {
                    continue;
                }
                if (ReportKompas.objectsAssemblyKompas.Count!=0)
                {
                    FileInfo fi = new FileInfo(ReportKompas.objectsAssemblyKompas[rowIndex].FullName);
                    path = fi.DirectoryName;
                }                    
            }
            Process.Start("explorer.exe", path);                   
        }
    }
}
