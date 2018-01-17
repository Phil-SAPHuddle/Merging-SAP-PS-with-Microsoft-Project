namespace ProjectAddIn2
{
    partial class frmSAPData
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblStatustext = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblProjectDef = new System.Windows.Forms.Label();
            this.sProjectDef = new System.Windows.Forms.TextBox();
            this.lstBoxProjectDef = new System.Windows.Forms.ListBox();
            this.lblProjectDefAttributes = new System.Windows.Forms.Label();
            this.tabTaskHier = new System.Windows.Forms.TabPage();
            this.dataGrid13 = new System.Windows.Forms.DataGridView();
            this.tabMsg = new System.Windows.Forms.TabPage();
            this.dataGrid5 = new System.Windows.Forms.DataGridView();
            this.tabWBSIn = new System.Windows.Forms.TabPage();
            this.dataGrid4 = new System.Windows.Forms.DataGridView();
            this.tabWBSHier = new System.Windows.Forms.TabPage();
            this.dataGrid3 = new System.Windows.Forms.DataGridView();
            this.tabNtwkActy = new System.Windows.Forms.TabPage();
            this.dataGrid2 = new System.Windows.Forms.DataGridView();
            this.tabWBS = new System.Windows.Forms.TabPage();
            this.dataGrid1 = new System.Windows.Forms.DataGridView();
            this.tabCtrl = new System.Windows.Forms.TabControl();
            this.statusStrip1.SuspendLayout();
            this.tabTaskHier.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid13)).BeginInit();
            this.tabMsg.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid5)).BeginInit();
            this.tabWBSIn.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid4)).BeginInit();
            this.tabWBSHier.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid3)).BeginInit();
            this.tabNtwkActy.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid2)).BeginInit();
            this.tabWBS.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
            this.tabCtrl.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblStatustext});
            this.statusStrip1.Location = new System.Drawing.Point(0, 567);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 12, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1127, 23);
            this.statusStrip1.TabIndex = 6;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lblStatustext
            // 
            this.lblStatustext.Name = "lblStatustext";
            this.lblStatustext.Size = new System.Drawing.Size(49, 18);
            this.lblStatustext.Text = "Status";
            // 
            // lblProjectDef
            // 
            this.lblProjectDef.AutoSize = true;
            this.lblProjectDef.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProjectDef.Location = new System.Drawing.Point(7, 28);
            this.lblProjectDef.Name = "lblProjectDef";
            this.lblProjectDef.Size = new System.Drawing.Size(98, 18);
            this.lblProjectDef.TabIndex = 10;
            this.lblProjectDef.Text = "Project Def:";
            // 
            // sProjectDef
            // 
            this.sProjectDef.Location = new System.Drawing.Point(111, 24);
            this.sProjectDef.Name = "sProjectDef";
            this.sProjectDef.Size = new System.Drawing.Size(223, 22);
            this.sProjectDef.TabIndex = 11;
            // 
            // lstBoxProjectDef
            // 
            this.lstBoxProjectDef.FormattingEnabled = true;
            this.lstBoxProjectDef.ItemHeight = 16;
            this.lstBoxProjectDef.Location = new System.Drawing.Point(12, 118);
            this.lstBoxProjectDef.Name = "lstBoxProjectDef";
            this.lstBoxProjectDef.Size = new System.Drawing.Size(209, 436);
            this.lstBoxProjectDef.TabIndex = 14;
            // 
            // lblProjectDefAttributes
            // 
            this.lblProjectDefAttributes.AutoSize = true;
            this.lblProjectDefAttributes.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProjectDefAttributes.Location = new System.Drawing.Point(12, 98);
            this.lblProjectDefAttributes.Name = "lblProjectDefAttributes";
            this.lblProjectDefAttributes.Size = new System.Drawing.Size(146, 17);
            this.lblProjectDefAttributes.TabIndex = 15;
            this.lblProjectDefAttributes.Text = "Project Def Attributes";
            // 
            // tabTaskHier
            // 
            this.tabTaskHier.Controls.Add(this.dataGrid13);
            this.tabTaskHier.Location = new System.Drawing.Point(4, 25);
            this.tabTaskHier.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabTaskHier.Name = "tabTaskHier";
            this.tabTaskHier.Size = new System.Drawing.Size(880, 438);
            this.tabTaskHier.TabIndex = 12;
            this.tabTaskHier.Text = "Task Hier";
            this.tabTaskHier.UseVisualStyleBackColor = true;
            // 
            // dataGrid13
            // 
            this.dataGrid13.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid13.Location = new System.Drawing.Point(12, 14);
            this.dataGrid13.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGrid13.Name = "dataGrid13";
            this.dataGrid13.RowTemplate.Height = 28;
            this.dataGrid13.Size = new System.Drawing.Size(853, 401);
            this.dataGrid13.TabIndex = 0;
            // 
            // tabMsg
            // 
            this.tabMsg.Controls.Add(this.dataGrid5);
            this.tabMsg.Location = new System.Drawing.Point(4, 25);
            this.tabMsg.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabMsg.Name = "tabMsg";
            this.tabMsg.Size = new System.Drawing.Size(880, 438);
            this.tabMsg.TabIndex = 4;
            this.tabMsg.Text = "Msg";
            this.tabMsg.UseVisualStyleBackColor = true;
            // 
            // dataGrid5
            // 
            this.dataGrid5.AllowUserToAddRows = false;
            this.dataGrid5.AllowUserToDeleteRows = false;
            this.dataGrid5.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.dataGrid5.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGrid5.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGrid5.ColumnHeadersHeight = 12;
            this.dataGrid5.Location = new System.Drawing.Point(4, 3);
            this.dataGrid5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGrid5.Name = "dataGrid5";
            this.dataGrid5.RowTemplate.Height = 28;
            this.dataGrid5.Size = new System.Drawing.Size(861, 428);
            this.dataGrid5.TabIndex = 0;
            // 
            // tabWBSIn
            // 
            this.tabWBSIn.Controls.Add(this.dataGrid4);
            this.tabWBSIn.Location = new System.Drawing.Point(4, 25);
            this.tabWBSIn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabWBSIn.Name = "tabWBSIn";
            this.tabWBSIn.Size = new System.Drawing.Size(880, 438);
            this.tabWBSIn.TabIndex = 3;
            this.tabWBSIn.Text = "WBSMLST";
            this.tabWBSIn.UseVisualStyleBackColor = true;
            // 
            // dataGrid4
            // 
            this.dataGrid4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid4.Location = new System.Drawing.Point(4, 3);
            this.dataGrid4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGrid4.Name = "dataGrid4";
            this.dataGrid4.RowTemplate.Height = 28;
            this.dataGrid4.Size = new System.Drawing.Size(858, 412);
            this.dataGrid4.TabIndex = 0;
            // 
            // tabWBSHier
            // 
            this.tabWBSHier.Controls.Add(this.dataGrid3);
            this.tabWBSHier.Location = new System.Drawing.Point(4, 25);
            this.tabWBSHier.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabWBSHier.Name = "tabWBSHier";
            this.tabWBSHier.Size = new System.Drawing.Size(880, 438);
            this.tabWBSHier.TabIndex = 2;
            this.tabWBSHier.Text = "WBSHier";
            this.tabWBSHier.UseVisualStyleBackColor = true;
            // 
            // dataGrid3
            // 
            this.dataGrid3.AllowUserToAddRows = false;
            this.dataGrid3.AllowUserToDeleteRows = false;
            this.dataGrid3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid3.Location = new System.Drawing.Point(13, 11);
            this.dataGrid3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGrid3.Name = "dataGrid3";
            this.dataGrid3.RowTemplate.Height = 28;
            this.dataGrid3.Size = new System.Drawing.Size(846, 409);
            this.dataGrid3.TabIndex = 0;
            // 
            // tabNtwkActy
            // 
            this.tabNtwkActy.Controls.Add(this.dataGrid2);
            this.tabNtwkActy.Location = new System.Drawing.Point(4, 25);
            this.tabNtwkActy.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabNtwkActy.Name = "tabNtwkActy";
            this.tabNtwkActy.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabNtwkActy.Size = new System.Drawing.Size(880, 438);
            this.tabNtwkActy.TabIndex = 1;
            this.tabNtwkActy.Text = "Ntwk Acty";
            this.tabNtwkActy.UseVisualStyleBackColor = true;
            // 
            // dataGrid2
            // 
            this.dataGrid2.AllowUserToOrderColumns = true;
            this.dataGrid2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid2.Location = new System.Drawing.Point(6, 14);
            this.dataGrid2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGrid2.Name = "dataGrid2";
            this.dataGrid2.RowTemplate.Height = 28;
            this.dataGrid2.Size = new System.Drawing.Size(846, 393);
            this.dataGrid2.TabIndex = 0;
            // 
            // tabWBS
            // 
            this.tabWBS.Controls.Add(this.dataGrid1);
            this.tabWBS.Location = new System.Drawing.Point(4, 25);
            this.tabWBS.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabWBS.Name = "tabWBS";
            this.tabWBS.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabWBS.Size = new System.Drawing.Size(880, 438);
            this.tabWBS.TabIndex = 0;
            this.tabWBS.Text = "WBS";
            this.tabWBS.UseVisualStyleBackColor = true;
            // 
            // dataGrid1
            // 
            this.dataGrid1.AllowUserToAddRows = false;
            this.dataGrid1.AllowUserToDeleteRows = false;
            this.dataGrid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGrid1.Location = new System.Drawing.Point(5, 16);
            this.dataGrid1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGrid1.Name = "dataGrid1";
            this.dataGrid1.RowTemplate.Height = 28;
            this.dataGrid1.Size = new System.Drawing.Size(869, 426);
            this.dataGrid1.TabIndex = 0;
            // 
            // tabCtrl
            // 
            this.tabCtrl.Controls.Add(this.tabWBS);
            this.tabCtrl.Controls.Add(this.tabNtwkActy);
            this.tabCtrl.Controls.Add(this.tabWBSHier);
            this.tabCtrl.Controls.Add(this.tabWBSIn);
            this.tabCtrl.Controls.Add(this.tabMsg);
            this.tabCtrl.Controls.Add(this.tabTaskHier);
            this.tabCtrl.Location = new System.Drawing.Point(227, 98);
            this.tabCtrl.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabCtrl.Name = "tabCtrl";
            this.tabCtrl.SelectedIndex = 0;
            this.tabCtrl.Size = new System.Drawing.Size(888, 467);
            this.tabCtrl.TabIndex = 0;
            // 
            // frmSAPData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1127, 590);
            this.Controls.Add(this.lblProjectDefAttributes);
            this.Controls.Add(this.lstBoxProjectDef);
            this.Controls.Add(this.sProjectDef);
            this.Controls.Add(this.lblProjectDef);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.tabCtrl);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "frmSAPData";
            this.Text = "SAPData";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.tabTaskHier.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid13)).EndInit();
            this.tabMsg.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid5)).EndInit();
            this.tabWBSIn.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid4)).EndInit();
            this.tabWBSHier.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid3)).EndInit();
            this.tabNtwkActy.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid2)).EndInit();
            this.tabWBS.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
            this.tabCtrl.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatustext;
        private System.Windows.Forms.Label lblProjectDef;
        private System.Windows.Forms.TextBox sProjectDef;
        private System.Windows.Forms.ListBox lstBoxProjectDef;
        private System.Windows.Forms.Label lblProjectDefAttributes;
        private System.Windows.Forms.TabPage tabTaskHier;
        private System.Windows.Forms.DataGridView dataGrid13;
        private System.Windows.Forms.TabPage tabMsg;
        private System.Windows.Forms.DataGridView dataGrid5;
        private System.Windows.Forms.TabPage tabWBSIn;
        private System.Windows.Forms.DataGridView dataGrid4;
        private System.Windows.Forms.TabPage tabWBSHier;
        private System.Windows.Forms.DataGridView dataGrid3;
        private System.Windows.Forms.TabPage tabNtwkActy;
        private System.Windows.Forms.DataGridView dataGrid2;
        private System.Windows.Forms.TabPage tabWBS;
        private System.Windows.Forms.DataGridView dataGrid1;
        private System.Windows.Forms.TabControl tabCtrl;
    }
}