using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

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
            XmlDocument doc = new XmlDocument();

            // Создаем корневой элемент
            XmlElement root = doc.CreateElement("Settings");
            doc.AppendChild(root);

            // Создаем элемент для первого TextBox
            XmlElement textBox1Element = doc.CreateElement("TextBox");
            textBox1Element.SetAttribute("Name", "CodeEquip");
            textBox1Element.InnerText = PathSettingsEquipmenttextBox.Text;
            root.AppendChild(textBox1Element);

            // Создаем элемент для второго TextBox
            XmlElement textBox2Element = doc.CreateElement("TextBox");
            textBox2Element.SetAttribute("Name", "CodeMaterial");
            textBox2Element.InnerText = PathSettingsMaterialtextBox.Text;
            root.AppendChild(textBox2Element);

            // Сохраняем XML в файл
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            doc.Save(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + "Settings.xml");
            //Properties.Settings.Default.PathSettingsEquipmenttextBox = PathSettingsEquipmenttextBox.Text;
            //Properties.Settings.Default.PathSettingsMaterialtextBox = PathSettingsMaterialtextBox.Text;
            //Properties.Settings.Default.Save();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\"))+@"\"+"CodeEquip.xml");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + "CodeMaterial.xml");
        }

    }
}
