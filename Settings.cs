using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportKompas
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void Equipment_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.PathSettingsEquipmenttextBox = PathSettingsEquipmenttextBox.Text;
            Properties.Settings.Default.PathSettingsMaterialtextBox = PathSettingsMaterialtextBox.Text;
            Properties.Settings.Default.Save();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\"))+@"\"+"CodeEquip.xml");
            //MessageBox.Show(tempPath.Remove(tempPath.LastIndexOf(@"\")));
            //Process.Start("explorer.exe", Properties.Settings.Default.PathSettingsEquipmenttextBox);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + "CodeMaterial.xml");
        }

    }
}
