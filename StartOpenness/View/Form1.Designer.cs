using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Reflection;
using System.IO;

namespace EPLAN_TIA
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txtSelect4 = new System.Windows.Forms.Label();
            this.txt_Status = new System.Windows.Forms.TextBox();
            this.btn_Start = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtStatusLabel = new System.Windows.Forms.Label();
            this.version = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txt_Path4 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.txtSelect3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_Path2 = new System.Windows.Forms.Button();
            this.btn_Path3 = new System.Windows.Forms.Button();
            this.txt_Path3 = new System.Windows.Forms.TextBox();
            this.txtSelect2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_Path2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSelect1 = new System.Windows.Forms.Label();
            this.btn_Path1 = new System.Windows.Forms.Button();
            this.txt_Path1 = new System.Windows.Forms.TextBox();
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
            // txtSelect4
            // 
            this.txtSelect4.AutoSize = true;
            this.txtSelect4.Location = new System.Drawing.Point(37, 243);
            this.txtSelect4.Name = "txtSelect4";
            this.txtSelect4.Size = new System.Drawing.Size(190, 13);
            this.txtSelect4.TabIndex = 44;
            this.txtSelect4.Text = "Select folder to save the data backups";
            // 
            // txt_Status
            // 
            this.txt_Status.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txt_Status.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(228)))), ((int)(((byte)(229)))), ((int)(((byte)(230)))));
            this.txt_Status.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_Status.Location = new System.Drawing.Point(88, 23);
            this.txt_Status.Multiline = true;
            this.txt_Status.Name = "txt_Status";
            this.txt_Status.ReadOnly = true;
            this.txt_Status.Size = new System.Drawing.Size(321, 48);
            this.txt_Status.TabIndex = 7;
            // 
            // btn_Start
            // 
            this.btn_Start.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(25)))), ((int)(((byte)(70)))));
            this.btn_Start.Enabled = false;
            this.btn_Start.Font = new System.Drawing.Font("Arial", 9.2F, System.Drawing.FontStyle.Bold);
            this.btn_Start.ForeColor = System.Drawing.Color.White;
            this.btn_Start.Location = new System.Drawing.Point(167, 311);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(89, 36);
            this.btn_Start.TabIndex = 17;
            this.btn_Start.Text = "Start";
            this.btn_Start.UseVisualStyleBackColor = false;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(228)))), ((int)(((byte)(229)))), ((int)(((byte)(230)))));
            this.panel1.Controls.Add(this.txtStatusLabel);
            this.panel1.Controls.Add(this.version);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.txt_Status);
            this.panel1.Location = new System.Drawing.Point(0, 373);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(431, 124);
            this.panel1.TabIndex = 34;
            // 
            // txtStatusLabel
            // 
            this.txtStatusLabel.AutoSize = true;
            this.txtStatusLabel.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtStatusLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(85)))), ((int)(((byte)(97)))));
            this.txtStatusLabel.Location = new System.Drawing.Point(26, 23);
            this.txtStatusLabel.Name = "txtStatusLabel";
            this.txtStatusLabel.Size = new System.Drawing.Size(44, 15);
            this.txtStatusLabel.TabIndex = 31;
            this.txtStatusLabel.Text = "Status";
            // 
            // version
            // 
            this.version.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.version.AutoSize = true;
            this.version.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.version.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(85)))), ((int)(((byte)(97)))));
            this.version.Location = new System.Drawing.Point(344, 93);
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
            this.label8.Location = new System.Drawing.Point(26, 104);
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
            this.panel2.Controls.Add(this.txt_Path4);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.txtSelect3);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.btn_Path2);
            this.panel2.Controls.Add(this.btn_Path3);
            this.panel2.Controls.Add(this.txtSelect4);
            this.panel2.Controls.Add(this.txt_Path3);
            this.panel2.Controls.Add(this.txtSelect2);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.txt_Path2);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.txtSelect1);
            this.panel2.Controls.Add(this.btn_Path1);
            this.panel2.Controls.Add(this.txt_Path1);
            this.panel2.Controls.Add(this.btn_Start);
            this.panel2.Location = new System.Drawing.Point(0, 31);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(431, 360);
            this.panel2.TabIndex = 26;
            // 
            // txt_Path4
            // 
            this.txt_Path4.Location = new System.Drawing.Point(37, 184);
            this.txt_Path4.Name = "txt_Path4";
            this.txt_Path4.Size = new System.Drawing.Size(361, 20);
            this.txt_Path4.TabIndex = 51;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(156)))), ((int)(((byte)(163)))));
            this.button1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(311, 210);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(89, 30);
            this.button1.TabIndex = 50;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtSelect3
            // 
            this.txtSelect3.AccessibleDescription = "";
            this.txtSelect3.AutoSize = true;
            this.txtSelect3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtSelect3.Location = new System.Drawing.Point(38, 167);
            this.txtSelect3.Name = "txtSelect3";
            this.txtSelect3.Size = new System.Drawing.Size(130, 13);
            this.txtSelect3.TabIndex = 49;
            this.txtSelect3.Text = "Excel  file export from TIA ";
            this.txtSelect3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
            // btn_Path2
            // 
            this.btn_Path2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(156)))), ((int)(((byte)(163)))));
            this.btn_Path2.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Path2.ForeColor = System.Drawing.Color.White;
            this.btn_Path2.Location = new System.Drawing.Point(312, 291);
            this.btn_Path2.Name = "btn_Path2";
            this.btn_Path2.Size = new System.Drawing.Size(89, 30);
            this.btn_Path2.TabIndex = 46;
            this.btn_Path2.Text = "Browse";
            this.btn_Path2.UseVisualStyleBackColor = false;
            this.btn_Path2.Click += new System.EventHandler(this.btn_Path2_Click);
            // 
            // btn_Path3
            // 
            this.btn_Path3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(156)))), ((int)(((byte)(163)))));
            this.btn_Path3.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Path3.ForeColor = System.Drawing.Color.White;
            this.btn_Path3.Location = new System.Drawing.Point(309, 136);
            this.btn_Path3.Name = "btn_Path3";
            this.btn_Path3.Size = new System.Drawing.Size(89, 30);
            this.btn_Path3.TabIndex = 45;
            this.btn_Path3.Text = "Browse";
            this.btn_Path3.UseVisualStyleBackColor = false;
            this.btn_Path3.Click += new System.EventHandler(this.btn_Path3_Click);
            // 
            // txt_Path3
            // 
            this.txt_Path3.Location = new System.Drawing.Point(35, 110);
            this.txt_Path3.Name = "txt_Path3";
            this.txt_Path3.Size = new System.Drawing.Size(361, 20);
            this.txt_Path3.TabIndex = 1;
            // 
            // txtSelect2
            // 
            this.txtSelect2.AutoSize = true;
            this.txtSelect2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtSelect2.Location = new System.Drawing.Point(33, 94);
            this.txtSelect2.Name = "txtSelect2";
            this.txtSelect2.Size = new System.Drawing.Size(203, 13);
            this.txtSelect2.TabIndex = 42;
            this.txtSelect2.Text = "Excel File with DataConnections (EPLAN)";
            this.txtSelect2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
            // txt_Path2
            // 
            this.txt_Path2.BackColor = System.Drawing.Color.White;
            this.txt_Path2.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Path2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(85)))), ((int)(((byte)(97)))));
            this.txt_Path2.Location = new System.Drawing.Point(37, 262);
            this.txt_Path2.Name = "txt_Path2";
            this.txt_Path2.Size = new System.Drawing.Size(364, 23);
            this.txt_Path2.TabIndex = 38;
            this.txt_Path2.TextChanged += new System.EventHandler(this.txt_Path2_TextChanged);
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
            // txtSelect1
            // 
            this.txtSelect1.AutoSize = true;
            this.txtSelect1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtSelect1.Location = new System.Drawing.Point(35, 20);
            this.txtSelect1.Name = "txtSelect1";
            this.txtSelect1.Size = new System.Drawing.Size(148, 13);
            this.txtSelect1.TabIndex = 35;
            this.txtSelect1.Text = " Excel File PLC Data (EPLAN)";
            this.txtSelect1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_Path1
            // 
            this.btn_Path1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(148)))), ((int)(((byte)(156)))), ((int)(((byte)(163)))));
            this.btn_Path1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Path1.ForeColor = System.Drawing.Color.White;
            this.btn_Path1.Location = new System.Drawing.Point(311, 65);
            this.btn_Path1.Name = "btn_Path1";
            this.btn_Path1.Size = new System.Drawing.Size(89, 30);
            this.btn_Path1.TabIndex = 34;
            this.btn_Path1.Text = "Browse";
            this.btn_Path1.UseVisualStyleBackColor = false;
            this.btn_Path1.Click += new System.EventHandler(this.btn_Path_Click);
            // 
            // txt_Path1
            // 
            this.txt_Path1.BackColor = System.Drawing.Color.White;
            this.txt_Path1.Font = new System.Drawing.Font("Arial", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Path1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(85)))), ((int)(((byte)(97)))));
            this.txt_Path1.Location = new System.Drawing.Point(36, 36);
            this.txt_Path1.Name = "txt_Path1";
            this.txt_Path1.Size = new System.Drawing.Size(364, 23);
            this.txt_Path1.TabIndex = 33;
            this.txt_Path1.TextChanged += new System.EventHandler(this.txt_Path_TextChanged);
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
            this.panel3.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel3_MouseDown);
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
            this.btn_DE.Click += new System.EventHandler(this.btn_DE_Click);
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
            this.btn_EN.Click += new System.EventHandler(this.btn_EN_Click);
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
            this.button_Minimize.Click += new System.EventHandler(this.button_Minimize_Click);
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
            this.button_Maximize.Click += new System.EventHandler(this.button_Maximize_Click);
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
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
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
            // Form1
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
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Openness_Start";
            this.Load += new System.EventHandler(this.Form1_Load);
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
        private System.Windows.Forms.TextBox txt_Status;
        private System.Windows.Forms.Button btn_Start;
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
        private System.Windows.Forms.Label txtStatusLabel;
        private System.Windows.Forms.Button btn_Path1;
        private System.Windows.Forms.TextBox txt_Path1;
        private System.Windows.Forms.Label txtSelect1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_Path2;
        private System.Windows.Forms.TextBox txt_Path3;
        private System.Windows.Forms.Label txtSelect2;
        private System.Windows.Forms.Button btn_Path2;
        private System.Windows.Forms.Button btn_Path3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label txtSelect3;
        private System.Windows.Forms.TextBox txt_Path4;
        private System.Windows.Forms.Label txtSelect4;
    }
}

