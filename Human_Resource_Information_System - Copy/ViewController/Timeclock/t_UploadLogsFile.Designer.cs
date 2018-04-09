namespace Human_Resource_Information_System
{
    partial class t_UploadLogsFile
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.pbar = new System.Windows.Forms.ProgressBar();
            this.btn_browse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btn_upload = new System.Windows.Forms.Button();
            this.dgvl_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_userid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvl_status = new System.Windows.Forms.DataGridViewTextBoxColumn();
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
            this.pnl_side.Size = new System.Drawing.Size(431, 496);
            this.pnl_side.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dgv_list);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(428, 496);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Upload Time Log History";
            // 
            // dgv_list
            // 
            this.dgv_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_list.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dgvl_date,
            this.dgvl_time,
            this.dgvl_userid,
            this.dgvl_status});
            this.dgv_list.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_list.Location = new System.Drawing.Point(3, 17);
            this.dgv_list.Name = "dgv_list";
            this.dgv_list.RowHeadersWidth = 25;
            this.dgv_list.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_list.Size = new System.Drawing.Size(422, 476);
            this.dgv_list.TabIndex = 0;
            this.dgv_list.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_list_CellContentClick_1);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.pbar);
            this.groupBox2.Controls.Add(this.btn_browse);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.textBox1);
            this.groupBox2.Controls.Add(this.btn_upload);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(431, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(571, 189);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Upload File";
            // 
            // pbar
            // 
            this.pbar.Location = new System.Drawing.Point(153, 160);
            this.pbar.Name = "pbar";
            this.pbar.Size = new System.Drawing.Size(314, 11);
            this.pbar.TabIndex = 4;
            // 
            // btn_browse
            // 
            this.btn_browse.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_browse.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_browse.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_browse.Location = new System.Drawing.Point(153, 77);
            this.btn_browse.Name = "btn_browse";
            this.btn_browse.Size = new System.Drawing.Size(154, 59);
            this.btn_browse.TabIndex = 3;
            this.btn_browse.Text = "Browse";
            this.btn_browse.UseVisualStyleBackColor = false;
            this.btn_browse.Click += new System.EventHandler(this.btn_browse_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "File Path";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(82, 28);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(455, 26);
            this.textBox1.TabIndex = 1;
            // 
            // btn_upload
            // 
            this.btn_upload.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_upload.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_upload.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_upload.Location = new System.Drawing.Point(313, 77);
            this.btn_upload.Name = "btn_upload";
            this.btn_upload.Size = new System.Drawing.Size(154, 59);
            this.btn_upload.TabIndex = 0;
            this.btn_upload.Text = "Upload";
            this.btn_upload.UseVisualStyleBackColor = false;
            this.btn_upload.Click += new System.EventHandler(this.btn_upload_Click);
            // 
            // dgvl_date
            // 
            this.dgvl_date.DataPropertyName = "work_date";
            this.dgvl_date.HeaderText = "Date";
            this.dgvl_date.Name = "dgvl_date";
            this.dgvl_date.ReadOnly = true;
            this.dgvl_date.Width = 77;
            // 
            // dgvl_time
            // 
            this.dgvl_time.DataPropertyName = "time_log";
            this.dgvl_time.HeaderText = "Time";
            this.dgvl_time.Name = "dgvl_time";
            this.dgvl_time.ReadOnly = true;
            this.dgvl_time.Width = 55;
            // 
            // dgvl_userid
            // 
            this.dgvl_userid.DataPropertyName = "empid";
            this.dgvl_userid.HeaderText = "Employee ID";
            this.dgvl_userid.Name = "dgvl_userid";
            this.dgvl_userid.ReadOnly = true;
            this.dgvl_userid.Width = 163;
            // 
            // dgvl_status
            // 
            this.dgvl_status.DataPropertyName = "status";
            this.dgvl_status.HeaderText = "Status";
            this.dgvl_status.Name = "dgvl_status";
            this.dgvl_status.ReadOnly = true;
            // 
            // t_UploadLogsFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1002, 496);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.pnl_side);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "t_UploadLogsFile";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Upload Time Logs File";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.t_UploadLogsFile_Load);
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
        private System.Windows.Forms.DataGridView dgv_list;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_browse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btn_upload;
        private System.Windows.Forms.ProgressBar pbar;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_date;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_time;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_userid;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvl_status;
    }
}