
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
            this.toolStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripButton4 = new System.Windows.Forms.ToolStripButton();
            this.Service = new System.Windows.Forms.ToolStripDropDownButton();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.collapseAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.expandAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.Export = new System.Windows.Forms.ToolStripDropDownButton();
            this.xmlBuhtaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.xmlTurboToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonPaint = new System.Windows.Forms.ToolStripButton();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.сохранитьВExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьВXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьВCsvToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьПапкуСОтчетомToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.MenuItemSelected = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьДиректориюСФайломToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.treePanel = new System.Windows.Forms.Panel();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton4,
            this.Service,
            this.Export,
            this.toolStripButton1,
            this.toolStripButtonPaint,
            this.toolStripLabel1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 399);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(884, 22);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton4
            // 
            this.toolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton4.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton4.Image")));
            this.toolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton4.Name = "toolStripButton4";
            this.toolStripButton4.Size = new System.Drawing.Size(112, 20);
            this.toolStripButton4.Text = "Прочитать сборку";
            this.toolStripButton4.Click += new System.EventHandler(this.toolStripButton4_Click);
            // 
            // Service
            // 
            this.Service.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Service.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem,
            this.collapseAllToolStripMenuItem,
            this.expandAllToolStripMenuItem,
            this.loadReportToolStripMenuItem});
            this.Service.Image = ((System.Drawing.Image)(resources.GetObject("Service.Image")));
            this.Service.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Service.Name = "Service";
            this.Service.Size = new System.Drawing.Size(60, 20);
            this.Service.Text = "Сервис";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(156, 22);
            this.settingsToolStripMenuItem.Text = "Настроить";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.SettingsToolStripMenuItem_Click);
            // 
            // collapseAllToolStripMenuItem
            // 
            this.collapseAllToolStripMenuItem.Name = "collapseAllToolStripMenuItem";
            this.collapseAllToolStripMenuItem.Size = new System.Drawing.Size(156, 22);
            this.collapseAllToolStripMenuItem.Text = "Свернуть всё";
            this.collapseAllToolStripMenuItem.Click += new System.EventHandler(this.CollapseAllToolStripMenuItem_Click);
            // 
            // expandAllToolStripMenuItem
            // 
            this.expandAllToolStripMenuItem.Name = "expandAllToolStripMenuItem";
            this.expandAllToolStripMenuItem.Size = new System.Drawing.Size(156, 22);
            this.expandAllToolStripMenuItem.Text = "Развернуть всё";
            this.expandAllToolStripMenuItem.Click += new System.EventHandler(this.ExpandAllToolStripMenuItem_Click);
            //
            // loadReportToolStripMenuItem
            //
            this.loadReportToolStripMenuItem.Name = "loadReportToolStripMenuItem";
            this.loadReportToolStripMenuItem.Size = new System.Drawing.Size(156, 22);
            this.loadReportToolStripMenuItem.Text = "Загрузить отчет";
            this.loadReportToolStripMenuItem.Click += new System.EventHandler(this.LoadReportToolStripMenuItem_Click);
            //
            // Export
            // 
            this.Export.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Export.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.xmlBuhtaToolStripMenuItem,
            this.xmlTurboToolStripMenuItem,
            this.excelToolStripMenuItem});
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(65, 20);
            this.Export.Text = "Экспорт";
            // 
            // xmlBuhtaToolStripMenuItem
            // 
            this.xmlBuhtaToolStripMenuItem.Name = "xmlBuhtaToolStripMenuItem";
            this.xmlBuhtaToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.xmlBuhtaToolStripMenuItem.Text = "Xml Buhta";
            this.xmlBuhtaToolStripMenuItem.Click += new System.EventHandler(this.StripButtonXML_Click);
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.excelToolStripMenuItem.Text = "Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.StripButtonExcel_Click);
            // 
            // xmlTurboToolStripMenuItem
            // 
            this.xmlTurboToolStripMenuItem.Name = "xmlTurboToolStripMenuItem";
            this.xmlTurboToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.xmlTurboToolStripMenuItem.Text = "Xml Turbo";
            this.xmlTurboToolStripMenuItem.Click += new System.EventHandler(this.XmlTurboMenuItem_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(151, 20);
            this.toolStripButton1.Text = "Открыть папку с отчетом";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripOpenExplorer_Click);
            // 
            // toolStripButtonPaint
            // 
            this.toolStripButtonPaint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonPaint.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonPaint.Image")));
            this.toolStripButtonPaint.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonPaint.Name = "toolStripButtonPaint";
            this.toolStripButtonPaint.Size = new System.Drawing.Size(23, 20);
            this.toolStripButtonPaint.Text = "toolStripButton2";
            this.toolStripButtonPaint.ToolTipText = "Назначить покрытие";
            this.toolStripButtonPaint.Click += new System.EventHandler(this.toolStripButtonPaint_Click);
            //
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(89, 17);
            this.toolStripLabel1.Text = "Строка статуса";
            // 
            // сохранитьВExcelToolStripMenuItem
            // 
            this.сохранитьВExcelToolStripMenuItem.Name = "сохранитьВExcelToolStripMenuItem";
            this.сохранитьВExcelToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // сохранитьВXMLToolStripMenuItem
            // 
            this.сохранитьВXMLToolStripMenuItem.Name = "сохранитьВXMLToolStripMenuItem";
            this.сохранитьВXMLToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // сохранитьВCsvToolStripMenuItem
            // 
            this.сохранитьВCsvToolStripMenuItem.Name = "сохранитьВCsvToolStripMenuItem";
            this.сохранитьВCsvToolStripMenuItem.Size = new System.Drawing.Size(214, 22);
            // 
            // открытьПапкуСОтчетомToolStripMenuItem
            // 
            this.открытьПапкуСОтчетомToolStripMenuItem.Name = "открытьПапкуСОтчетомToolStripMenuItem";
            this.открытьПапкуСОтчетомToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // MenuItemSelected
            // 
            this.MenuItemSelected.Name = "MenuItemSelected";
            this.MenuItemSelected.Size = new System.Drawing.Size(67, 22);
            // 
            // открытьДиректориюСФайломToolStripMenuItem
            // 
            this.открытьДиректориюСФайломToolStripMenuItem.Name = "открытьДиректориюСФайломToolStripMenuItem";
            this.открытьДиректориюСФайломToolStripMenuItem.Size = new System.Drawing.Size(67, 22);
            //
            // treePanel
            //
            this.treePanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treePanel.Location = new System.Drawing.Point(0, 0);
            this.treePanel.Name = "treePanel";
            this.treePanel.Size = new System.Drawing.Size(884, 399);
            this.treePanel.TabIndex = 1;
            //
            // ReportKompas
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 421);
            this.Controls.Add(this.treePanel);
            this.Controls.Add(this.toolStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(900, 460);
            this.Name = "ReportKompas";
            this.Text = "ReportKompas";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip toolStrip1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem MenuItemSelected;
        private System.Windows.Forms.ToolStripMenuItem открытьДиректориюСФайломToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьВExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem открытьПапкуСОтчетомToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьВCsvToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьВXMLToolStripMenuItem;
        private System.Windows.Forms.ToolStripDropDownButton Service;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem collapseAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem expandAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButton4;
        private System.Windows.Forms.ToolStripStatusLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripDropDownButton Export;
        private System.Windows.Forms.ToolStripMenuItem xmlBuhtaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem xmlTurboToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripButton toolStripButtonPaint;
        private System.Windows.Forms.Panel treePanel;
    }
}

