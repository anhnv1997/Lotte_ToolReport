
namespace ToolReport
{
    partial class ucMail
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgvMail = new System.Windows.Forms.DataGridView();
            this._ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.tsbADD = new System.Windows.Forms.ToolStripButton();
            this.tsbEDIT = new System.Windows.Forms.ToolStripButton();
            this.tsbDELETE = new System.Windows.Forms.ToolStripButton();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMail)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvMail
            // 
            this.dgvMail.AllowUserToAddRows = false;
            this.dgvMail.AllowUserToDeleteRows = false;
            this.dgvMail.AllowUserToResizeColumns = false;
            this.dgvMail.AllowUserToResizeRows = false;
            this.dgvMail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvMail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvMail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvMail.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dgvMail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMail.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this._ID,
            this.Column2});
            this.dgvMail.Location = new System.Drawing.Point(0, 28);
            this.dgvMail.MultiSelect = false;
            this.dgvMail.Name = "dgvMail";
            this.dgvMail.RowHeadersVisible = false;
            this.dgvMail.RowTemplate.Height = 25;
            this.dgvMail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvMail.Size = new System.Drawing.Size(450, 319);
            this.dgvMail.TabIndex = 0;
            // 
            // _ID
            // 
            this._ID.HeaderText = "ID";
            this._ID.Name = "_ID";
            this._ID.Width = 43;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "MailAddress";
            this.Column2.Name = "Column2";
            this.Column2.Width = 97;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbADD,
            this.tsbEDIT,
            this.tsbDELETE});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(450, 25);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // tsbADD
            // 
            this.tsbADD.Image = global::ToolReport.Properties.Resources.add;
            this.tsbADD.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbADD.Name = "tsbADD";
            this.tsbADD.Size = new System.Drawing.Size(51, 22);
            this.tsbADD.Text = "ADD";
            this.tsbADD.Click += new System.EventHandler(this.tsbADD_Click);
            // 
            // tsbEDIT
            // 
            this.tsbEDIT.Image = global::ToolReport.Properties.Resources.edit;
            this.tsbEDIT.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbEDIT.Name = "tsbEDIT";
            this.tsbEDIT.Size = new System.Drawing.Size(47, 22);
            this.tsbEDIT.Text = "Edit";
            this.tsbEDIT.Click += new System.EventHandler(this.tsbEDIT_Click);
            // 
            // tsbDELETE
            // 
            this.tsbDELETE.Image = global::ToolReport.Properties.Resources.delete;
            this.tsbDELETE.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbDELETE.Name = "tsbDELETE";
            this.tsbDELETE.Size = new System.Drawing.Size(60, 22);
            this.tsbDELETE.Text = "Delete";
            this.tsbDELETE.Click += new System.EventHandler(this.tsbDELETE_Click);
            // 
            // ucMail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.dgvMail);
            this.Name = "ucMail";
            this.Size = new System.Drawing.Size(450, 347);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMail)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvMail;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton tsbADD;
        private System.Windows.Forms.ToolStripButton tsbEDIT;
        private System.Windows.Forms.ToolStripButton tsbDELETE;
        private System.Windows.Forms.DataGridViewTextBoxColumn _ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
    }
}
