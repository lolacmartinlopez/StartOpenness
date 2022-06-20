using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Reflection;
using System.IO;

namespace EPLAN_TIA
{
    partial class Form2
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components2 = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components2 != null))
            {
                components2.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            ""}, -1);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.panel1 = new System.Windows.Forms.Panel();
            this.version = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.EPLAN_Data = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.TIA_Data = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btn_DE = new System.Windows.Forms.Button();
            this.btn_EN = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.button_Minimize = new System.Windows.Forms.Button();
            this.button_Maximize = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(228)))), ((int)(((byte)(229)))), ((int)(((byte)(230)))));
            this.panel1.Controls.Add(this.version);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Location = new System.Drawing.Point(0, 397);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(431, 100);
            this.panel1.TabIndex = 34;
            // 
            // version
            // 
            this.version.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.version.AutoSize = true;
            this.version.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.version.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(85)))), ((int)(((byte)(97)))));
            this.version.Location = new System.Drawing.Point(354, 74);
            this.version.Name = "version";
            this.version.Size = new System.Drawing.Size(55, 19);
            this.version.TabIndex = 29;
            this.version.Text = "V1.0.0";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arial", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(85)))), ((int)(((byte)(97)))));
            this.label8.Location = new System.Drawing.Point(21, 77);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(282, 13);
            this.label8.TabIndex = 27;
            this.label8.Text = "© 2022 EDAG Production Solutions GmbH Co. KG. ";
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.listView1);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Location = new System.Drawing.Point(0, 31);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(431, 438);
            this.panel2.TabIndex = 26;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(26, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(306, 18);
            this.label1.TabIndex = 49;
            this.label1.Text = "The next Data are different in both documents";
            // 
            // listView1
            // 
            this.listView1.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.listView1.AllowColumnReorder = true;
            this.listView1.AllowDrop = true;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.EPLAN_Data,
            this.TIA_Data});
            this.listView1.FullRowSelect = true;
            this.listView1.GridLines = true;
            this.listView1.HideSelection = false;
            listViewItem1.UseItemStyleForSubItems = false;
            this.listView1.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1});
            this.listView1.Location = new System.Drawing.Point(24, 62);
            this.listView1.Name = "listView1";
            this.listView1.OwnerDraw = true;
            this.listView1.Size = new System.Drawing.Size(385, 361);
            this.listView1.TabIndex = 48;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.UseWaitCursor = true;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // EPLAN_Data
            // 
            this.EPLAN_Data.Text = "EPLAN Data";
            this.EPLAN_Data.Width = 192;
            // 
            // TIA_Data
            // 
            this.TIA_Data.Text = "TIA Data";
            this.TIA_Data.Width = 316;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(39, 184);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 47;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(33, 248);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(0, 13);
            this.label5.TabIndex = 41;
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(35, 110);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 37;
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(25)))), ((int)(((byte)(70)))));
            this.panel3.Controls.Add(this.btn_DE);
            this.panel3.Controls.Add(this.btn_EN);
            this.panel3.Controls.Add(this.label7);
            this.panel3.Controls.Add(this.button_Minimize);
            this.panel3.Controls.Add(this.button_Maximize);
            this.panel3.Controls.Add(this.button_Exit);
            this.panel3.Controls.Add(this.pictureBox1);
            this.panel3.Location = new System.Drawing.Point(0, -1);
            this.panel3.Margin = new System.Windows.Forms.Padding(2);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(431, 43);
            this.panel3.TabIndex = 35;
            // 
            // btn_DE
            // 
            this.btn_DE.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_DE.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(164)))), ((int)(((byte)(13)))), ((int)(((byte)(48)))));
            this.btn_DE.FlatAppearance.BorderSize = 0;
            this.btn_DE.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_DE.Font = new System.Drawing.Font("Arial", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_DE.ForeColor = System.Drawing.Color.White;
            this.btn_DE.Location = new System.Drawing.Point(256, -2);
            this.btn_DE.Margin = new System.Windows.Forms.Padding(2);
            this.btn_DE.Name = "btn_DE";
            this.btn_DE.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btn_DE.Size = new System.Drawing.Size(44, 50);
            this.btn_DE.TabIndex = 39;
            this.btn_DE.Text = "DE";
            this.btn_DE.UseVisualStyleBackColor = true;
            // 
            // btn_EN
            // 
            this.btn_EN.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_EN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(164)))), ((int)(((byte)(13)))), ((int)(((byte)(48)))));
            this.btn_EN.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(164)))), ((int)(((byte)(13)))), ((int)(((byte)(48)))));
            this.btn_EN.FlatAppearance.BorderSize = 0;
            this.btn_EN.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_EN.Font = new System.Drawing.Font("Arial", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_EN.ForeColor = System.Drawing.Color.White;
            this.btn_EN.Location = new System.Drawing.Point(212, -2);
            this.btn_EN.Margin = new System.Windows.Forms.Padding(2);
            this.btn_EN.Name = "btn_EN";
            this.btn_EN.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btn_EN.Size = new System.Drawing.Size(44, 50);
            this.btn_EN.TabIndex = 36;
            this.btn_EN.Text = "EN";
            this.btn_EN.UseVisualStyleBackColor = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(103, 14);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(86, 16);
            this.label7.TabIndex = 34;
            this.label7.Text = "EPLAN->TIA";
            // 
            // button_Minimize
            // 
            this.button_Minimize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Minimize.FlatAppearance.BorderSize = 0;
            this.button_Minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Minimize.Font = new System.Drawing.Font("Arial Narrow", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Minimize.ForeColor = System.Drawing.Color.White;
            this.button_Minimize.Location = new System.Drawing.Point(298, -1);
            this.button_Minimize.Margin = new System.Windows.Forms.Padding(2);
            this.button_Minimize.Name = "button_Minimize";
            this.button_Minimize.Size = new System.Drawing.Size(40, 53);
            this.button_Minimize.TabIndex = 6;
            this.button_Minimize.Text = "–";
            this.button_Minimize.UseVisualStyleBackColor = true;
            // 
            // button_Maximize
            // 
            this.button_Maximize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Maximize.FlatAppearance.BorderSize = 0;
            this.button_Maximize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Maximize.Font = new System.Drawing.Font("Arial", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Maximize.ForeColor = System.Drawing.Color.White;
            this.button_Maximize.Location = new System.Drawing.Point(344, -3);
            this.button_Maximize.Margin = new System.Windows.Forms.Padding(2);
            this.button_Maximize.Name = "button_Maximize";
            this.button_Maximize.Size = new System.Drawing.Size(40, 50);
            this.button_Maximize.TabIndex = 5;
            this.button_Maximize.Text = "□";
            this.button_Maximize.UseVisualStyleBackColor = true;
            // 
            // button_Exit
            // 
            this.button_Exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Exit.FlatAppearance.BorderSize = 0;
            this.button_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Exit.Font = new System.Drawing.Font("Arial Narrow", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.ForeColor = System.Drawing.Color.White;
            this.button_Exit.Location = new System.Drawing.Point(390, -3);
            this.button_Exit.Margin = new System.Windows.Forms.Padding(2);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.button_Exit.Size = new System.Drawing.Size(40, 50);
            this.button_Exit.TabIndex = 0;
            this.button_Exit.Text = "x";
            this.button_Exit.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::StartOpenness.Properties.Resources.white_logo_png;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(98, 43);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(431, 496);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Openness_Start";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }



        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button_Minimize;
        private System.Windows.Forms.Button button_Maximize;
        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btn_EN;
        private System.Windows.Forms.Button btn_DE;
        private System.Windows.Forms.Label version;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ColumnHeader EPLAN_Data;
        public System.Windows.Forms.ColumnHeader TIA_Data;
        public System.Windows.Forms.ListView listView1;
    }
}

