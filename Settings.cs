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
            LoadSet();
        }

        void LoadSet()
        {
            XmlDocument xmlDoc = new XmlDocument();
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            xmlDoc.Load(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + "Settings.xml");
            var textBoxesNodes = xmlDoc.SelectNodes("/Settings/TextBox");
            foreach (XmlNode node in textBoxesNodes)
            {
                string name = node.Attributes["Name"].Value;
                string value = node.InnerText;

                if (name == "CodeEquip") // добавляем установку PathSettingsEquipmenttextBox.Text
                    PathSettingsEquipmenttextBox.Text = value;
                else if (name == "CodeMaterial") // добавляем установку PathSettingsEquipmenttextBox.Text
                    PathSettingsMaterialtextBox.Text = value;
            }
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(PathSettingsEquipmenttextBox.Text + @"\"+"CodeEquip.xml");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(PathSettingsMaterialtextBox.Text + @"\" + "CodeMaterial.xml");
        }

    }
}
