
namespace ReportKompas
{
    partial class ReportKompas
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportKompas));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.сохранитьВExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьВCsvToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьПапкуСОтчетомToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripButton();
            this.Service = new System.Windows.Forms.ToolStripDropDownButton();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.MenuItemSelected = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьДиректориюСФайломToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripDropDownButton1,
            this.toolStripButton3,
            this.Service,
            this.toolStripButton2});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1043, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(119, 22);
            this.toolStripButton1.Text = "Отобразить данные";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButtonShowData_Click);
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.сохранитьВExcelToolStripMenuItem,
            this.сохранитьВXMLToolStripMenuItem,
            this.сохранитьВCsvToolStripMenuItem,
            this.открытьПапкуСОтчетомToolStripMenuItem});
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(52, 22);
            this.toolStripDropDownButton1.Text = "Отчет";
            // 
            // сохранитьВExcelToolStripMenuItem
            // 
            this.сохранитьВExcelToolStripMenuItem.Name = "сохранитьВExcelToolStripMenuItem";
            this.сохранитьВExcelToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.сохранитьВExcelToolStripMenuItem.Text = "Сохранить в Excel";
            this.сохранитьВExcelToolStripMenuItem.Click += new System.EventHandler(this.SaveExcel_ToolStripMenuItem_Click);
            // 
            // сохранитьВXMLToolStripMenuItem
            // 
            this.сохранитьВXMLToolStripMenuItem.Name = "сохранитьВXMLToolStripMenuItem";
            this.сохранитьВXMLToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.сохранитьВXMLToolStripMenuItem.Text = "Сохранить в XML";
            this.сохранитьВXMLToolStripMenuItem.Click += new System.EventHandler(this.SaveXML_ToolStripMenuItem_Click);
            // 
            // сохранитьВCsvToolStripMenuItem
            // 
            this.сохранитьВCsvToolStripMenuItem.Name = "сохранитьВCsvToolStripMenuItem";
            this.сохранитьВCsvToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.сохранитьВCsvToolStripMenuItem.Text = "Сохранить в csv";
            this.сохранитьВCsvToolStripMenuItem.Click += new System.EventHandler(this.SaveCSV_ToolStripMenuItem_Click);
            // 
            // открытьПапкуСОтчетомToolStripMenuItem
            // 
            this.открытьПапкуСОтчетомToolStripMenuItem.Name = "открытьПапкуСОтчетомToolStripMenuItem";
            this.открытьПапкуСОтчетомToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            this.открытьПапкуСОтчетомToolStripMenuItem.Text = "Открыть папку с отчетом";
            this.открытьПапкуСОтчетомToolStripMenuItem.Click += new System.EventHandler(this.OpenExplorer_ToolStripMenuItem_Click);
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton3.Image")));
            this.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(194, 22);
            this.toolStripButton3.Text = "Показать пропущенные позиции";
            this.toolStripButton3.Click += new System.EventHandler(this.ShowLostParts_toolStripButton_Click);
            // 
            // Service
            // 
            this.Service.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Service.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem});
            this.Service.Image = ((System.Drawing.Image)(resources.GetObject("Service.Image")));
            this.Service.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Service.Name = "Service";
            this.Service.Size = new System.Drawing.Size(60, 22);
            this.Service.Text = "Сервис";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(132, 22);
            this.settingsToolStripMenuItem.Text = "Настроить";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.SettingsToolStripMenuItem_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(59, 22);
            this.toolStripButton2.Text = "Колонки";
            this.toolStripButton2.Click += new System.EventHandler(this.toolStripButton2_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.ContextMenuStrip = this.contextMenuStrip1;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 25);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1043, 425);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuItemSelected,
            this.открытьДиректориюСФайломToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(241, 48);
            this.contextMenuStrip1.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStrip1_Opening);
            // 
            // MenuItemSelected
            // 
            this.MenuItemSelected.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.MenuItemSelected.Name = "MenuItemSelected";
            this.MenuItemSelected.Size = new System.Drawing.Size(240, 22);
            this.MenuItemSelected.Text = "&Открыть выбранное в Компас";
            this.MenuItemSelected.Click += new System.EventHandler(this.MenuItemOpenInKompas_Click);
            // 
            // открытьДиректориюСФайломToolStripMenuItem
            // 
            this.открытьДиректориюСФайломToolStripMenuItem.Name = "открытьДиректориюСФайломToolStripMenuItem";
            this.открытьДиректориюСФайломToolStripMenuItem.Size = new System.Drawing.Size(240, 22);
            this.открытьДиректориюСФайломToolStripMenuItem.Text = "&Открыть в проводнике";
            this.открытьДиректориюСФайломToolStripMenuItem.Click += new System.EventHandler(this.OpenExplorer_MenuItem_Click);
            // 
            // ReportKompas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1043, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.toolStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ReportKompas";
            this.Text = "ReportKompas";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem MenuItemSelected;
        private System.Windows.Forms.ToolStripMenuItem открытьДиректориюСФайломToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButton3;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripMenuItem сохранитьВExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem открытьПапкуСОтчетомToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьВCsvToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьВXMLToolStripMenuItem;
        private System.Windows.Forms.ToolStripDropDownButton Service;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButton2;
    }
}

