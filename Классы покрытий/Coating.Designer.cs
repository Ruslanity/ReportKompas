namespace ReportKompas
{
    partial class Coating
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
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnAssignCoating = new System.Windows.Forms.ToolStripButton();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.colDesignation = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCoating = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCoverageArea = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colIsPainted = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnAssignCoating});
            this.toolStrip1.Location = new System.Drawing.Point(0, 425);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(643, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // btnAssignCoating
            // 
            this.btnAssignCoating.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnAssignCoating.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnAssignCoating.Name = "btnAssignCoating";
            this.btnAssignCoating.Size = new System.Drawing.Size(126, 22);
            this.btnAssignCoating.Text = "Назначить покрытие";
            this.btnAssignCoating.Click += new System.EventHandler(this.btnAssignCoating_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colDesignation,
            this.colName,
            this.colCoating,
            this.colCoverageArea,
            this.colIsPainted});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(643, 425);
            this.dataGridView1.TabIndex = 1;
            // 
            // colDesignation
            // 
            this.colDesignation.HeaderText = "Обозначение";
            this.colDesignation.Name = "colDesignation";
            this.colDesignation.ReadOnly = true;
            this.colDesignation.Width = 200;
            // 
            // colName
            // 
            this.colName.HeaderText = "Наименование";
            this.colName.Name = "colName";
            this.colName.ReadOnly = true;
            this.colName.Width = 200;
            // 
            // colCoating
            // 
            this.colCoating.HeaderText = "Покрытие";
            this.colCoating.Name = "colCoating";
            this.colCoating.Width = 80;
            // 
            // colCoverageArea
            // 
            this.colCoverageArea.HeaderText = "Площадь покрытия";
            this.colCoverageArea.Name = "colCoverageArea";
            this.colCoverageArea.Width = 80;
            // 
            // colIsPainted
            // 
            this.colIsPainted.HeaderText = "IsPainted";
            this.colIsPainted.Name = "colIsPainted";
            this.colIsPainted.Width = 80;
            // 
            // Coating
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(643, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.toolStrip1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Coating";
            this.Text = "Покрытия";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnAssignCoating;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDesignation;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCoating;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCoverageArea;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colIsPainted;
    }
}
