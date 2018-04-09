namespace Human_Resource_Information_System
{
    partial class p_GeneratePayroll
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
            this.pnl_side = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dgv_list = new System.Windows.Forms.DataGridView();
            this.dgvl_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_payroll = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_employee = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_userid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.pbar = new System.Windows.Forms.ProgressBar();
            this.btn_generate = new System.Windows.Forms.Button();
            this.bg_worker = new System.ComponentModel.BackgroundWorker();
            this.pnl_side.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnl_side
            // 
            this.pnl_side.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_side.Controls.Add(this.groupBox1);
            this.pnl_side.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnl_side.Location = new System.Drawing.Point(0, 0);
            this.pnl_side.Name = "pnl_side";
            this.pnl_side.Size = new System.Drawing.Size(431, 442);
            this.pnl_side.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dgv_list);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(428, 442);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Generated DTR History";
            // 
            // dgv_list
            // 
            this.dgv_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_list.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dgvl_date,
            this.dgvl_time,
            this.dgvl_payroll,
            this.dgvl_employee,
            this.dgvl_userid});
            this.dgv_list.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_list.Location = new System.Drawing.Point(3, 17);
            this.dgv_list.Name = "dgv_list";
            this.dgv_list.ReadOnly = true;
            this.dgv_list.RowHeadersWidth = 25;
            this.dgv_list.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_list.Size = new System.Drawing.Size(422, 422);
            this.dgv_list.TabIndex = 1;
            // 
            // dgvl_date
            // 
            this.dgvl_date.HeaderText = "Date Generated";
            this.dgvl_date.Name = "dgvl_date";
            this.dgvl_date.ReadOnly = true;
            this.dgvl_date.Width = 77;
            // 
            // dgvl_time
            // 
            this.dgvl_time.HeaderText = "Time Generated";
            this.dgvl_time.Name = "dgvl_time";
            this.dgvl_time.ReadOnly = true;
            this.dgvl_time.Width = 55;
            // 
            // dgvl_payroll
            // 
            this.dgvl_payroll.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dgvl_payroll.HeaderText = "Payroll Period";
            this.dgvl_payroll.Name = "dgvl_payroll";
            this.dgvl_payroll.ReadOnly = true;
            // 
            // dgvl_employee
            // 
            this.dgvl_employee.HeaderText = "Employee";
            this.dgvl_employee.Name = "dgvl_employee";
            this.dgvl_employee.ReadOnly = true;
            // 
            // dgvl_userid
            // 
            this.dgvl_userid.HeaderText = "User ID";
            this.dgvl_userid.Name = "dgvl_userid";
            this.dgvl_userid.ReadOnly = true;
            this.dgvl_userid.Width = 77;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkBox3);
            this.groupBox2.Controls.Add(this.checkBox2);
            this.groupBox2.Controls.Add(this.checkBox1);
            this.groupBox2.Controls.Add(this.pbar);
            this.groupBox2.Controls.Add(this.btn_generate);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(431, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(677, 238);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Generate Payroll";
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(296, 54);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(164, 19);
            this.checkBox3.TabIndex = 4;
            this.checkBox3.Text = "Include Other Deductions";
            this.checkBox3.UseVisualStyleBackColor = true;
            this.checkBox3.Visible = false;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(78, 54);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(151, 19);
            this.checkBox2.TabIndex = 3;
            this.checkBox2.Text = "Include Other Earnings";
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.Visible = false;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(523, 54);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(97, 19);
            this.checkBox1.TabIndex = 2;
            this.checkBox1.Text = "Include Loan";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            // 
            // pbar
            // 
            this.pbar.Location = new System.Drawing.Point(78, 100);
            this.pbar.Name = "pbar";
            this.pbar.Size = new System.Drawing.Size(542, 40);
            this.pbar.TabIndex = 1;
            // 
            // btn_generate
            // 
            this.btn_generate.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_generate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_generate.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_generate.Location = new System.Drawing.Point(270, 160);
            this.btn_generate.Name = "btn_generate";
            this.btn_generate.Size = new System.Drawing.Size(154, 59);
            this.btn_generate.TabIndex = 0;
            this.btn_generate.Text = "Generate";
            this.btn_generate.UseVisualStyleBackColor = false;
            this.btn_generate.Click += new System.EventHandler(this.btn_generate_Click);
            // 
            // bg_worker
            // 
            this.bg_worker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bg_worker_DoWork);
            // 
            // p_GeneratePayroll
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1108, 442);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.pnl_side);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "p_GeneratePayroll";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = " ";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.p_GeneratePayroll_Load);
            this.pnl_side.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_side;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_generate;
        private System.Windows.Forms.ProgressBar pbar;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.ComponentModel.BackgroundWorker bg_worker;
        private System.Windows.Forms.DataGridView dgv_list;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_date;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_time;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_payroll;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_employee;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_userid;
    }
}