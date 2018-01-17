
using System.Data;

namespace ProjectAddIn2
{
    partial class frmProjectSelection
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
            this.dataGridPrjDefs = new System.Windows.Forms.DataGridView();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblServer = new System.Windows.Forms.Label();
            this.lblServerID = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridPrjDefs)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridPrjDefs
            // 
            this.dataGridPrjDefs.AllowUserToAddRows = false;
            this.dataGridPrjDefs.AllowUserToDeleteRows = false;
            this.dataGridPrjDefs.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridPrjDefs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridPrjDefs.Location = new System.Drawing.Point(27, 74);
            this.dataGridPrjDefs.Name = "dataGridPrjDefs";
            this.dataGridPrjDefs.RowTemplate.Height = 24;
            this.dataGridPrjDefs.Size = new System.Drawing.Size(538, 648);
            this.dataGridPrjDefs.TabIndex = 0;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(193, 743);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 40);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(292, 743);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 40);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblServer
            // 
            this.lblServer.AutoSize = true;
            this.lblServer.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblServer.Location = new System.Drawing.Point(24, 25);
            this.lblServer.Name = "lblServer";
            this.lblServer.Size = new System.Drawing.Size(99, 18);
            this.lblServer.TabIndex = 3;
            this.lblServer.Text = "SAP Server:";
            // 
            // lblServerID
            // 
            this.lblServerID.AutoSize = true;
            this.lblServerID.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblServerID.Location = new System.Drawing.Point(129, 25);
            this.lblServerID.Name = "lblServerID";
            this.lblServerID.Size = new System.Drawing.Size(13, 18);
            this.lblServerID.TabIndex = 4;
            this.lblServerID.Text = " ";
            // 
            // frmProjectSelection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 795);
            this.Controls.Add(this.lblServerID);
            this.Controls.Add(this.lblServer);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.dataGridPrjDefs);
            this.Name = "frmProjectSelection";
            this.Text = "SAP Project Definitions";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridPrjDefs)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridPrjDefs;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblServer;
        private System.Windows.Forms.Label lblServerID;
    }
}