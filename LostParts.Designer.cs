
namespace ReportKompas
{
    partial class LostParts
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ToolStripMenuItemOpenInKompas = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьВПроводникеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.ContextMenuStrip = this.contextMenuStrip1;
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.Location = new System.Drawing.Point(0, 0);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(800, 450);
            this.dataGridView2.TabIndex = 0;
            this.dataGridView2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView2_MouseDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItemOpenInKompas,
            this.открытьВПроводникеToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(200, 70);
            // 
            // ToolStripMenuItemOpenInKompas
            // 
            this.ToolStripMenuItemOpenInKompas.Name = "ToolStripMenuItemOpenInKompas";
            this.ToolStripMenuItemOpenInKompas.Size = new System.Drawing.Size(199, 22);
            this.ToolStripMenuItemOpenInKompas.Text = "&Открыть в Компас";
            this.ToolStripMenuItemOpenInKompas.Click += new System.EventHandler(this.ToolStripMenuItemOpenInKompas_Click);
            // 
            // открытьВПроводникеToolStripMenuItem
            // 
            this.открытьВПроводникеToolStripMenuItem.Name = "открытьВПроводникеToolStripMenuItem";
            this.открытьВПроводникеToolStripMenuItem.Size = new System.Drawing.Size(199, 22);
            this.открытьВПроводникеToolStripMenuItem.Text = "&Открыть в проводнике";
            this.открытьВПроводникеToolStripMenuItem.Click += new System.EventHandler(this.ToolStripMenuItemOpenInExplorer_Click);
            // 
            // LostParts
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.dataGridView2);
            this.Name = "LostParts";
            this.Text = "LostParts";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemOpenInKompas;
        private System.Windows.Forms.ToolStripMenuItem открытьВПроводникеToolStripMenuItem;
    }
}