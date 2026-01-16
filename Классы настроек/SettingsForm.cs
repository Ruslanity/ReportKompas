using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ReportKompas
{
    class SettingsForm : Form
    {
        private Settings settings;

        private GroupBox groupBox;
        private TableLayoutPanel layoutPanel;
        private RowStyle rowStyle1;
        private GroupBox groupBox1;
        private TableLayoutPanel layoutPanel1;
        private Button buttonEquipment;
        private RowStyle rowStyle2;
        private GroupBox groupBox2;
        private TableLayoutPanel layoutPanel2;
        private Button buttonMaterial;
        private RowStyle rowStyle3;
        private GroupBox groupBox3;
        private TableLayoutPanel layoutPanel3;
        private Button buttonSpeedCut;
        private RowStyle rowStyle4;
        private GroupBox groupBox4;
        private TextBox txtEquipment;
        private TextBox txtMaterials;
        private TextBox txtSpeedCut;
        private CheckBox chkCalcLaserCutTime;

        public SettingsForm(Settings settings)
        {
            this.settings = settings;

            InitializeComponents();
            LoadSettings();
            this.Size = new Size(500, 270);
            MinimumSize = new Size(500, 270);
            Load += (sender, args) => OnSizeChanged(EventArgs.Empty);
            SizeChanged += (sender, args) =>
            {
                this.Padding = new Padding(3);
                this.MaximizeBox = false;
                this.MinimizeBox = false;

                rowStyle1.SizeType = SizeType.Percent;
                rowStyle1.Height = 33.3f;
                rowStyle2.SizeType = SizeType.Percent;
                rowStyle2.Height = 33.3f;
                rowStyle3.SizeType = SizeType.Percent;
                rowStyle3.Height = 33.3f;
                rowStyle4.SizeType = SizeType.Absolute;
                rowStyle4.Height = 50;
            };
            // В конструкторе или методе инициализации
            this.FormClosing += new FormClosingEventHandler(CloseSettings);
            this.Focus();
        }

        private void InitializeComponents()
        {
            Text = "Настройки";
            groupBox = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Пути к библиотекам"
            };
            Controls.Add(groupBox);
            layoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1
            };
            groupBox.Controls.Add(layoutPanel);

            #region Строчка СтИ
            rowStyle1 = new RowStyle();
            layoutPanel.RowStyles.Add(rowStyle1);
            groupBox1 = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Путь до словаря стандартных изделий"
            };
            layoutPanel.Controls.Add(groupBox1, 0, 0);
            layoutPanel1 = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2
            };
            groupBox1.Controls.Add(layoutPanel1);
            layoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // первая колонка
            layoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80)); // вторая колонка
            txtEquipment = new TextBox
            {
                Location = new Point(10, 20),
                Dock = DockStyle.Fill,
                Font = new Font(@"Microsoft Sans Serif", 8.25f, FontStyle.Regular),
                Multiline = true
            };
            layoutPanel1.Controls.Add(txtEquipment, 0, 0);
            buttonEquipment = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                Text = "Обзор"
            };
            buttonEquipment.Click += (sender, e) =>
            {
                if (System.IO.File.Exists(txtEquipment.Text))
                {
                    // Открытие файла по умолчанию
                    Process.Start(txtEquipment.Text);
                }
                else
                {
                    MessageBox.Show("Файл не найден по указанному пути.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }                
            };
            layoutPanel1.Controls.Add(buttonEquipment, 1, 0);
            #endregion

            #region Строчка материалов
            rowStyle2 = new RowStyle();
            layoutPanel.RowStyles.Add(rowStyle2);
            groupBox2 = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Путь до словаря материалов"
            };
            layoutPanel.Controls.Add(groupBox2, 0, 1);
            layoutPanel2 = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2
            };
            groupBox2.Controls.Add(layoutPanel2);
            layoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // первая колонка
            layoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80)); // вторая колонка
            txtMaterials = new TextBox
            {
                Location = new Point(10, 20),
                Dock = DockStyle.Fill,
                Font = new Font(@"Microsoft Sans Serif", 8.25f, FontStyle.Regular),
                Multiline = true
            };
            layoutPanel2.Controls.Add(txtMaterials, 0, 0);
            buttonMaterial = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                Text = "Обзор"
            };
            buttonMaterial.Click += (sender, e) =>
            {
                if (System.IO.File.Exists(txtMaterials.Text))
                {
                    // Открытие файла по умолчанию
                    Process.Start(txtMaterials.Text);
                }
                else
                {
                    MessageBox.Show("Файл не найден по указанному пути.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            layoutPanel2.Controls.Add(buttonMaterial, 1, 0);
            #endregion

            #region Строка настроек скорости резки
            rowStyle3 = new RowStyle();
            layoutPanel.RowStyles.Add(rowStyle3);
            groupBox3 = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Путь до настроек скорости резки"
            };
            layoutPanel.Controls.Add(groupBox3, 0, 2);
            layoutPanel3 = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2
            };
            groupBox3.Controls.Add(layoutPanel3);
            layoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // первая колонка
            layoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80)); // вторая колонка
            txtSpeedCut = new TextBox
            {
                Location = new Point(10, 20),
                Dock = DockStyle.Fill,
                Font = new Font(@"Microsoft Sans Serif", 8.25f, FontStyle.Regular),
                Multiline = true
            };
            layoutPanel3.Controls.Add(txtSpeedCut, 0, 0);
            buttonSpeedCut = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                Text = "Обзор"
            };
            buttonSpeedCut.Click += (sender, e) =>
            {
                if (System.IO.File.Exists(txtSpeedCut.Text))
                {
                    // Открытие файла по умолчанию
                    Process.Start(txtSpeedCut.Text);
                }
                else
                {
                    MessageBox.Show("Файл не найден по указанному пути.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            layoutPanel3.Controls.Add(buttonSpeedCut, 1, 0);
            #endregion

            #region Строка с прочее
            rowStyle4 = new RowStyle();
            layoutPanel.RowStyles.Add(rowStyle4);
            groupBox4 = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Прочие параметры"
            };
            layoutPanel.Controls.Add(groupBox4, 0, 3);
            chkCalcLaserCutTime = new CheckBox
            {
                Top = 20,
                Left = 10,
                Text = "Выполнять расчет времени резки и габаритных размеров",
                Size = new Size(330, 20)
            };
            groupBox4.Controls.Add(chkCalcLaserCutTime);
            #endregion

        }

        private void LoadSettings()
        {
            txtMaterials.Text = settings.PathDictionaryMaterials;
            txtEquipment.Text = settings.PathDictionaryEquipment;
            txtSpeedCut.Text = settings.PathDictionarySpeedCut;
            chkCalcLaserCutTime.Checked = settings.CalcLaserCutTime;
        }

        private void CloseSettings(object sender, EventArgs e)
        {
            settings.PathDictionaryMaterials = txtMaterials.Text;
            settings.PathDictionaryEquipment = txtEquipment.Text;
            settings.PathDictionarySpeedCut = txtSpeedCut.Text;
            settings.CalcLaserCutTime = chkCalcLaserCutTime.Checked;
            settings.Save(Settings.DefaultPathSettings);
        }
    }
}
