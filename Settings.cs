using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
            xmlDoc.Load(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + @"Settings\Settings.xml");
            var textBoxesNodes = xmlDoc.SelectNodes("/Settings/TextBox");
            foreach (XmlNode node in textBoxesNodes)
            {
                string name = node.Attributes["Name"].Value;
                string value = node.InnerText;

                if (name == "CodeEquip") // добавляем установку PathSettingsEquipmenttextBox.Text
                    Path_Dictionary_Equipment_textBox.Text = value;
                else if (name == "CodeMaterial") // добавляем установку PathSettingsEquipmenttextBox.Text
                    Path_Dictionary_Materials_textBox.Text = value;
                else if (name == "SpeedCut") // добавляем установку PathSettingsEquipmenttextBox.Text
                    Speed_Cut_textBox.Text = value;
            }
            // Обработка чекбокса
            var checkBoxNode = xmlDoc.SelectSingleNode("/Settings/CheckBox[@Name='Other_Param_Laser_Cut_checkBox']");
            if (checkBoxNode != null)
            {
                bool isChecked;
                if (bool.TryParse(checkBoxNode.InnerText, out isChecked))
                {
                    Other_Param_Laser_Cut_checkBox.Checked = isChecked;
                }
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
            textBox1Element.InnerText = Path_Dictionary_Equipment_textBox.Text;
            root.AppendChild(textBox1Element);

            // Создаем элемент для второго TextBox
            XmlElement textBox2Element = doc.CreateElement("TextBox");
            textBox2Element.SetAttribute("Name", "CodeMaterial");
            textBox2Element.InnerText = Path_Dictionary_Materials_textBox.Text;
            root.AppendChild(textBox2Element);

            // Создаем элемент для второго TextBox
            XmlElement textBox3Element = doc.CreateElement("TextBox");
            textBox3Element.SetAttribute("Name", "SpeedCut");
            textBox3Element.InnerText = Speed_Cut_textBox.Text;
            root.AppendChild(textBox3Element);

            // Создаем элемент для CheckBox
            XmlElement checkBoxElement = doc.CreateElement("CheckBox");
            checkBoxElement.SetAttribute("Name", "Other_Param_Laser_Cut_checkBox"); // Можно указать имя чекбокса
            checkBoxElement.InnerText = Other_Param_Laser_Cut_checkBox.Checked.ToString(); // сохраняем состояние как строку "True" или "False"
            root.AppendChild(checkBoxElement);

            // Сохраняем XML в файл            
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directoryPath = System.IO.Path.GetDirectoryName(tempPath);
            string filePath = System.IO.Path.Combine(directoryPath, @"Settings\Settings.xml");

            //using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read))
            //{
            //    doc.Save(fs);
            //}
            doc.Save(filePath);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(Path_Dictionary_Equipment_textBox.Text + @"\"+"CodeEquip.xml");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(Path_Dictionary_Materials_textBox.Text + @"\" + "CodeMaterial.xml");
        }

        private void Speed_Cut_button_Click(object sender, EventArgs e)
        {
            Process.Start(Speed_Cut_textBox.Text + @"\" + "SpeedCut.xml");
        }
    }
}
