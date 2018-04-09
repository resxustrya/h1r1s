namespace Human_Resource_Information_System
{
    partial class z_Notification
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pnl_sidebar = new System.Windows.Forms.Panel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cbo_type = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cbo_branch = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dtp_to = new System.Windows.Forms.DateTimePicker();
            this.dtp_frm = new System.Windows.Forms.DateTimePicker();
            this.cbo_module = new System.Windows.Forms.ComboBox();
            this.btn_new = new System.Windows.Forms.Button();
            this.pnl_main = new System.Windows.Forms.Panel();
            this.dgv_list = new System.Windows.Forms.DataGridView();
            this.t_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.t_time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.branch = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.module = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.username = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.notification_text = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnl_sidebar.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.pnl_main.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).BeginInit();
            this.SuspendLayout();
            // 
            // pnl_sidebar
            // 
            this.pnl_sidebar.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_sidebar.Controls.Add(this.groupBox4);
            this.pnl_sidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnl_sidebar.Location = new System.Drawing.Point(0, 0);
            this.pnl_sidebar.Name = "pnl_sidebar";
            this.pnl_sidebar.Size = new System.Drawing.Size(228, 500);
            this.pnl_sidebar.TabIndex = 45;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.cbo_type);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.cbo_branch);
            this.groupBox4.Controls.Add(this.label1);
            this.groupBox4.Controls.Add(this.dtp_to);
            this.groupBox4.Controls.Add(this.dtp_frm);
            this.groupBox4.Controls.Add(this.cbo_module);
            this.groupBox4.Controls.Add(this.btn_new);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.ForeColor = System.Drawing.SystemColors.Info;
            this.groupBox4.Location = new System.Drawing.Point(9, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(210, 311);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Main Option";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 139);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(81, 16);
            this.label5.TabIndex = 10;
            this.label5.Text = "Select Type";
            // 
            // cbo_type
            // 
            this.cbo_type.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_type.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbo_type.FormattingEnabled = true;
            this.cbo_type.Items.AddRange(new object[] {
            "All",
            "Reminders",
            "Notifications"});
            this.cbo_type.Location = new System.Drawing.Point(10, 158);
            this.cbo_type.Name = "cbo_type";
            this.cbo_type.Size = new System.Drawing.Size(184, 24);
            this.cbo_type.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 87);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(91, 16);
            this.label4.TabIndex = 8;
            this.label4.Text = "Select Branch";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 187);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(94, 16);
            this.label3.TabIndex = 6;
            this.label3.Text = "Select Module";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "Date To";
            // 
            // cbo_branch
            // 
            this.cbo_branch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_branch.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbo_branch.FormattingEnabled = true;
            this.cbo_branch.Location = new System.Drawing.Point(10, 106);
            this.cbo_branch.Name = "cbo_branch";
            this.cbo_branch.Size = new System.Drawing.Size(184, 24);
            this.cbo_branch.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 16);
            this.label1.TabIndex = 4;
            this.label1.Text = "Date From";
            // 
            // dtp_to
            // 
            this.dtp_to.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtp_to.Location = new System.Drawing.Point(84, 52);
            this.dtp_to.Name = "dtp_to";
            this.dtp_to.Size = new System.Drawing.Size(110, 22);
            this.dtp_to.TabIndex = 3;
            this.dtp_to.ValueChanged += new System.EventHandler(this.dtp_to_ValueChanged);
            // 
            // dtp_frm
            // 
            this.dtp_frm.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtp_frm.Location = new System.Drawing.Point(84, 22);
            this.dtp_frm.Name = "dtp_frm";
            this.dtp_frm.Size = new System.Drawing.Size(110, 22);
            this.dtp_frm.TabIndex = 2;
            this.dtp_frm.ValueChanged += new System.EventHandler(this.dtp_frm_ValueChanged);
            // 
            // cbo_module
            // 
            this.cbo_module.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_module.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbo_module.FormattingEnabled = true;
            this.cbo_module.Items.AddRange(new object[] {
            "Customer",
            "Supplier",
            "Journal Entry",
            "Posting To Ledger",
            "Statement of Account",
            "Purchase Orders",
            "Direct Purchases",
            "Customer Legder",
            "Supplier Ledger",
            "Aging"});
            this.cbo_module.Location = new System.Drawing.Point(10, 206);
            this.cbo_module.Name = "cbo_module";
            this.cbo_module.Size = new System.Drawing.Size(184, 24);
            this.cbo_module.TabIndex = 1;
            // 
            // btn_new
            // 
            this.btn_new.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_new.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_new.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_new.Location = new System.Drawing.Point(10, 242);
            this.btn_new.Name = "btn_new";
            this.btn_new.Size = new System.Drawing.Size(184, 49);
            this.btn_new.TabIndex = 0;
            this.btn_new.Text = "   View Notification";
            this.btn_new.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_new.UseVisualStyleBackColor = false;
            this.btn_new.Click += new System.EventHandler(this.btn_new_Click);
            // 
            // pnl_main
            // 
            this.pnl_main.Controls.Add(this.dgv_list);
            this.pnl_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnl_main.Location = new System.Drawing.Point(228, 0);
            this.pnl_main.Name = "pnl_main";
            this.pnl_main.Size = new System.Drawing.Size(562, 500);
            this.pnl_main.TabIndex = 46;
            // 
            // dgv_list
            // 
            this.dgv_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_list.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.t_date,
            this.t_time,
            this.branch,
            this.type,
            this.module,
            this.username,
            this.notification_text});
            this.dgv_list.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_list.Location = new System.Drawing.Point(0, 0);
            this.dgv_list.Name = "dgv_list";
            this.dgv_list.RowHeadersWidth = 20;
            this.dgv_list.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_list.Size = new System.Drawing.Size(562, 500);
            this.dgv_list.TabIndex = 0;
            // 
            // t_date
            // 
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.t_date.DefaultCellStyle = dataGridViewCellStyle1;
            this.t_date.HeaderText = "DATE";
            this.t_date.Name = "t_date";
            this.t_date.ReadOnly = true;
            this.t_date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // t_time
            // 
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.t_time.DefaultCellStyle = dataGridViewCellStyle2;
            this.t_time.HeaderText = "TIME";
            this.t_time.Name = "t_time";
            this.t_time.ReadOnly = true;
            this.t_time.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.t_time.Width = 77;
            // 
            // branch
            // 
            this.branch.HeaderText = "BRANCH";
            this.branch.Name = "branch";
            this.branch.ReadOnly = true;
            // 
            // type
            // 
            this.type.HeaderText = "TYPE";
            this.type.Name = "type";
            this.type.ReadOnly = true;
            this.type.Width = 50;
            // 
            // module
            // 
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.module.DefaultCellStyle = dataGridViewCellStyle3;
            this.module.HeaderText = "MODULE";
            this.module.Name = "module";
            this.module.ReadOnly = true;
            this.module.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.module.Width = 150;
            // 
            // username
            // 
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.username.DefaultCellStyle = dataGridViewCellStyle4;
            this.username.HeaderText = "USERNAME";
            this.username.Name = "username";
            this.username.ReadOnly = true;
            this.username.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // notification_text
            // 
            this.notification_text.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.notification_text.DefaultCellStyle = dataGridViewCellStyle5;
            this.notification_text.HeaderText = "MESSAGE";
            this.notification_text.Name = "notification_text";
            this.notification_text.ReadOnly = true;
            this.notification_text.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // z_Notification
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(790, 500);
            this.Controls.Add(this.pnl_main);
            this.Controls.Add(this.pnl_sidebar);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "z_Notification";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Notification";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.z_Notification_Load);
            this.pnl_sidebar.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.pnl_main.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_sidebar;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbo_type;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbo_branch;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtp_to;
        private System.Windows.Forms.DateTimePicker dtp_frm;
        private System.Windows.Forms.ComboBox cbo_module;
        private System.Windows.Forms.Button btn_new;
        private System.Windows.Forms.Panel pnl_main;
        private System.Windows.Forms.DataGridView dgv_list;
        private System.Windows.Forms.DataGridViewTextBoxColumn t_date;
        private System.Windows.Forms.DataGridViewTextBoxColumn t_time;
        private System.Windows.Forms.DataGridViewTextBoxColumn branch;
        private System.Windows.Forms.DataGridViewTextBoxColumn type;
        private System.Windows.Forms.DataGridViewTextBoxColumn module;
        private System.Windows.Forms.DataGridViewTextBoxColumn username;
        private System.Windows.Forms.DataGridViewTextBoxColumn notification_text;
    }
}