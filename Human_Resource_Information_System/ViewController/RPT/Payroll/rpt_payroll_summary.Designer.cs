namespace Human_Resource_Information_System
{
    partial class rpt_payroll_summary
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
            this.pnl_sidebar = new System.Windows.Forms.Panel();
            this.tbcntrl_option = new System.Windows.Forms.TabControl();
            this.tpg_opt_1 = new System.Windows.Forms.TabPage();
            this.panel5 = new System.Windows.Forms.Panel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btn_deletefile = new System.Windows.Forms.Button();
            this.btn_print = new System.Windows.Forms.Button();
            this.empname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pay_code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pay_period = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bgworker = new System.ComponentModel.BackgroundWorker();
            this.pnl_rpt_option_header = new System.Windows.Forms.Panel();
            this.pnl_rpt_option = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pic_loading = new System.Windows.Forms.PictureBox();
            this.btn_submit = new System.Windows.Forms.Button();
            this.grp_options = new System.Windows.Forms.GroupBox();
            this.cbo_payollperiod = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbo_employee = new System.Windows.Forms.ComboBox();
            this.lbl_cbo_1 = new System.Windows.Forms.Label();
            this.empid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ppid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tbcntrl_main = new System.Windows.Forms.TabControl();
            this.tpg_list = new System.Windows.Forms.TabPage();
            this.panel7 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgvl_payroll_summary = new System.Windows.Forms.DataGridView();
            this.rpt_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.filename = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.date_added = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.cbo_department = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.pnl_sidebar.SuspendLayout();
            this.tbcntrl_option.SuspendLayout();
            this.tpg_opt_1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.pnl_rpt_option_header.SuspendLayout();
            this.pnl_rpt_option.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_loading)).BeginInit();
            this.grp_options.SuspendLayout();
            this.tbcntrl_main.SuspendLayout();
            this.tpg_list.SuspendLayout();
            this.panel7.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvl_payroll_summary)).BeginInit();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnl_sidebar
            // 
            this.pnl_sidebar.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_sidebar.Controls.Add(this.tbcntrl_option);
            this.pnl_sidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnl_sidebar.Location = new System.Drawing.Point(0, 110);
            this.pnl_sidebar.Margin = new System.Windows.Forms.Padding(4);
            this.pnl_sidebar.Name = "pnl_sidebar";
            this.pnl_sidebar.Size = new System.Drawing.Size(263, 667);
            this.pnl_sidebar.TabIndex = 78;
            // 
            // tbcntrl_option
            // 
            this.tbcntrl_option.Alignment = System.Windows.Forms.TabAlignment.Bottom;
            this.tbcntrl_option.Controls.Add(this.tpg_opt_1);
            this.tbcntrl_option.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_option.Location = new System.Drawing.Point(0, 0);
            this.tbcntrl_option.Margin = new System.Windows.Forms.Padding(4);
            this.tbcntrl_option.Name = "tbcntrl_option";
            this.tbcntrl_option.SelectedIndex = 0;
            this.tbcntrl_option.Size = new System.Drawing.Size(263, 667);
            this.tbcntrl_option.TabIndex = 1;
            // 
            // tpg_opt_1
            // 
            this.tpg_opt_1.Controls.Add(this.panel5);
            this.tpg_opt_1.Location = new System.Drawing.Point(4, 4);
            this.tpg_opt_1.Margin = new System.Windows.Forms.Padding(4);
            this.tpg_opt_1.Name = "tpg_opt_1";
            this.tpg_opt_1.Size = new System.Drawing.Size(255, 638);
            this.tpg_opt_1.TabIndex = 2;
            this.tpg_opt_1.UseVisualStyleBackColor = true;
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.SteelBlue;
            this.panel5.Controls.Add(this.groupBox4);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(0, 0);
            this.panel5.Margin = new System.Windows.Forms.Padding(4);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(255, 638);
            this.panel5.TabIndex = 0;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btn_deletefile);
            this.groupBox4.Controls.Add(this.btn_print);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.ForeColor = System.Drawing.SystemColors.Info;
            this.groupBox4.Location = new System.Drawing.Point(0, 0);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox4.Size = new System.Drawing.Size(255, 601);
            this.groupBox4.TabIndex = 64;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Main Option";
            // 
            // btn_deletefile
            // 
            this.btn_deletefile.BackColor = System.Drawing.Color.Maroon;
            this.btn_deletefile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_deletefile.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_deletefile.Location = new System.Drawing.Point(11, 143);
            this.btn_deletefile.Margin = new System.Windows.Forms.Padding(4);
            this.btn_deletefile.Name = "btn_deletefile";
            this.btn_deletefile.Size = new System.Drawing.Size(208, 89);
            this.btn_deletefile.TabIndex = 8;
            this.btn_deletefile.Text = "Delete File";
            this.btn_deletefile.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_deletefile.UseVisualStyleBackColor = false;
            this.btn_deletefile.Click += new System.EventHandler(this.btn_deletefile_Click);
            // 
            // btn_print
            // 
            this.btn_print.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_print.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_print.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_print.Location = new System.Drawing.Point(11, 47);
            this.btn_print.Margin = new System.Windows.Forms.Padding(4);
            this.btn_print.Name = "btn_print";
            this.btn_print.Size = new System.Drawing.Size(208, 89);
            this.btn_print.TabIndex = 3;
            this.btn_print.Text = "View";
            this.btn_print.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_print.UseVisualStyleBackColor = false;
            this.btn_print.Click += new System.EventHandler(this.btn_print_Click);
            // 
            // empname
            // 
            this.empname.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.empname.HeaderText = "EMPLOYEE NAME";
            this.empname.Name = "empname";
            this.empname.ReadOnly = true;
            // 
            // pay_code
            // 
            this.pay_code.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.pay_code.HeaderText = "PAYROLL CODE";
            this.pay_code.Name = "pay_code";
            this.pay_code.ReadOnly = true;
            // 
            // pay_period
            // 
            this.pay_period.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.pay_period.HeaderText = "PAYROLL PERIOD";
            this.pay_period.Name = "pay_period";
            this.pay_period.ReadOnly = true;
            // 
            // bgworker
            // 
            this.bgworker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgworker_DoWork);
            // 
            // pnl_rpt_option_header
            // 
            this.pnl_rpt_option_header.Controls.Add(this.pnl_rpt_option);
            this.pnl_rpt_option_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_rpt_option_header.Location = new System.Drawing.Point(0, 0);
            this.pnl_rpt_option_header.Margin = new System.Windows.Forms.Padding(4);
            this.pnl_rpt_option_header.Name = "pnl_rpt_option_header";
            this.pnl_rpt_option_header.Size = new System.Drawing.Size(1604, 110);
            this.pnl_rpt_option_header.TabIndex = 77;
            // 
            // pnl_rpt_option
            // 
            this.pnl_rpt_option.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_rpt_option.Controls.Add(this.panel2);
            this.pnl_rpt_option.Controls.Add(this.grp_options);
            this.pnl_rpt_option.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_rpt_option.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnl_rpt_option.Location = new System.Drawing.Point(0, 0);
            this.pnl_rpt_option.Margin = new System.Windows.Forms.Padding(4);
            this.pnl_rpt_option.Name = "pnl_rpt_option";
            this.pnl_rpt_option.Size = new System.Drawing.Size(1604, 106);
            this.pnl_rpt_option.TabIndex = 8;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.pic_loading);
            this.panel2.Controls.Add(this.btn_submit);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(840, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(228, 106);
            this.panel2.TabIndex = 4;
            // 
            // pic_loading
            // 
            this.pic_loading.ErrorImage = global::Human_Resource_Information_System.Properties.Resources.spin;
            this.pic_loading.Image = global::Human_Resource_Information_System.Properties.Resources.spin;
            this.pic_loading.Location = new System.Drawing.Point(81, 73);
            this.pic_loading.Margin = new System.Windows.Forms.Padding(4);
            this.pic_loading.Name = "pic_loading";
            this.pic_loading.Size = new System.Drawing.Size(43, 32);
            this.pic_loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic_loading.TabIndex = 7;
            this.pic_loading.TabStop = false;
            // 
            // btn_submit
            // 
            this.btn_submit.BackColor = System.Drawing.Color.SaddleBrown;
            this.btn_submit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_submit.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_submit.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btn_submit.Location = new System.Drawing.Point(8, 12);
            this.btn_submit.Margin = new System.Windows.Forms.Padding(4);
            this.btn_submit.Name = "btn_submit";
            this.btn_submit.Size = new System.Drawing.Size(185, 53);
            this.btn_submit.TabIndex = 1;
            this.btn_submit.Text = "Generate Payroll Summary Report";
            this.btn_submit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_submit.UseVisualStyleBackColor = false;
            this.btn_submit.Click += new System.EventHandler(this.btn_submit_Click);
            // 
            // grp_options
            // 
            this.grp_options.Controls.Add(this.cbo_department);
            this.grp_options.Controls.Add(this.label2);
            this.grp_options.Controls.Add(this.cbo_payollperiod);
            this.grp_options.Controls.Add(this.label1);
            this.grp_options.Controls.Add(this.cbo_employee);
            this.grp_options.Controls.Add(this.lbl_cbo_1);
            this.grp_options.Dock = System.Windows.Forms.DockStyle.Left;
            this.grp_options.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grp_options.ForeColor = System.Drawing.SystemColors.Info;
            this.grp_options.Location = new System.Drawing.Point(0, 0);
            this.grp_options.Margin = new System.Windows.Forms.Padding(4);
            this.grp_options.Name = "grp_options";
            this.grp_options.Padding = new System.Windows.Forms.Padding(4);
            this.grp_options.Size = new System.Drawing.Size(840, 106);
            this.grp_options.TabIndex = 3;
            this.grp_options.TabStop = false;
            this.grp_options.Text = "Other Options";
            // 
            // cbo_payollperiod
            // 
            this.cbo_payollperiod.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_payollperiod.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_payollperiod.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_payollperiod.FormattingEnabled = true;
            this.cbo_payollperiod.Location = new System.Drawing.Point(153, 19);
            this.cbo_payollperiod.Margin = new System.Windows.Forms.Padding(4);
            this.cbo_payollperiod.Name = "cbo_payollperiod";
            this.cbo_payollperiod.Size = new System.Drawing.Size(455, 26);
            this.cbo_payollperiod.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 22);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(100, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "Payroll Period";
            // 
            // cbo_employee
            // 
            this.cbo_employee.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_employee.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_employee.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_employee.FormattingEnabled = true;
            this.cbo_employee.Location = new System.Drawing.Point(153, 44);
            this.cbo_employee.Margin = new System.Windows.Forms.Padding(4);
            this.cbo_employee.Name = "cbo_employee";
            this.cbo_employee.Size = new System.Drawing.Size(455, 26);
            this.cbo_employee.TabIndex = 1;
            // 
            // lbl_cbo_1
            // 
            this.lbl_cbo_1.AutoSize = true;
            this.lbl_cbo_1.Location = new System.Drawing.Point(9, 49);
            this.lbl_cbo_1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbl_cbo_1.Name = "lbl_cbo_1";
            this.lbl_cbo_1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_cbo_1.Size = new System.Drawing.Size(74, 18);
            this.lbl_cbo_1.TabIndex = 1;
            this.lbl_cbo_1.Text = "Employee";
            // 
            // empid
            // 
            this.empid.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.empid.HeaderText = "EMPLOYEE ID";
            this.empid.Name = "empid";
            this.empid.ReadOnly = true;
            // 
            // ppid
            // 
            this.ppid.HeaderText = "ppid";
            this.ppid.Name = "ppid";
            this.ppid.ReadOnly = true;
            this.ppid.Visible = false;
            // 
            // tbcntrl_main
            // 
            this.tbcntrl_main.Controls.Add(this.tpg_list);
            this.tbcntrl_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_main.Location = new System.Drawing.Point(263, 110);
            this.tbcntrl_main.Margin = new System.Windows.Forms.Padding(4);
            this.tbcntrl_main.Name = "tbcntrl_main";
            this.tbcntrl_main.SelectedIndex = 0;
            this.tbcntrl_main.Size = new System.Drawing.Size(1341, 667);
            this.tbcntrl_main.TabIndex = 79;
            // 
            // tpg_list
            // 
            this.tpg_list.Controls.Add(this.panel7);
            this.tpg_list.Location = new System.Drawing.Point(4, 25);
            this.tpg_list.Margin = new System.Windows.Forms.Padding(4);
            this.tpg_list.Name = "tpg_list";
            this.tpg_list.Padding = new System.Windows.Forms.Padding(4);
            this.tpg_list.Size = new System.Drawing.Size(1333, 638);
            this.tpg_list.TabIndex = 0;
            this.tpg_list.Text = "Payroll Summary PDF List";
            this.tpg_list.UseVisualStyleBackColor = true;
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.SystemColors.Info;
            this.panel7.Controls.Add(this.groupBox2);
            this.panel7.Controls.Add(this.panel4);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(4, 4);
            this.panel7.Margin = new System.Windows.Forms.Padding(4);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(1325, 630);
            this.panel7.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgvl_payroll_summary);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox2.Location = new System.Drawing.Point(0, 55);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(1325, 575);
            this.groupBox2.TabIndex = 58;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Payroll Period List";
            // 
            // dgvl_payroll_summary
            // 
            this.dgvl_payroll_summary.AllowUserToResizeRows = false;
            this.dgvl_payroll_summary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvl_payroll_summary.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.rpt_id,
            this.filename,
            this.date_added});
            this.dgvl_payroll_summary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvl_payroll_summary.Location = new System.Drawing.Point(4, 23);
            this.dgvl_payroll_summary.Margin = new System.Windows.Forms.Padding(4);
            this.dgvl_payroll_summary.Name = "dgvl_payroll_summary";
            this.dgvl_payroll_summary.ReadOnly = true;
            this.dgvl_payroll_summary.RowHeadersVisible = false;
            this.dgvl_payroll_summary.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvl_payroll_summary.Size = new System.Drawing.Size(1317, 548);
            this.dgvl_payroll_summary.TabIndex = 1;
            // 
            // rpt_id
            // 
            this.rpt_id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.rpt_id.HeaderText = "CODE";
            this.rpt_id.Name = "rpt_id";
            this.rpt_id.ReadOnly = true;
            // 
            // filename
            // 
            this.filename.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.filename.HeaderText = "PAYROLL SUMMARY PDF FILES";
            this.filename.Name = "filename";
            this.filename.ReadOnly = true;
            // 
            // date_added
            // 
            this.date_added.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.date_added.HeaderText = "DATE CREATED";
            this.date_added.Name = "date_added";
            this.date_added.ReadOnly = true;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.label5);
            this.panel4.Controls.Add(this.label19);
            this.panel4.Controls.Add(this.dateTimePicker1);
            this.panel4.Controls.Add(this.dateTimePicker2);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Margin = new System.Windows.Forms.Padding(4);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1325, 55);
            this.panel4.TabIndex = 59;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(311, 21);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(25, 17);
            this.label5.TabIndex = 70;
            this.label5.Text = "To";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(16, 21);
            this.label19.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(124, 17);
            this.label19.TabIndex = 69;
            this.label19.Text = "Transaction Dates";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(352, 15);
            this.dateTimePicker1.Margin = new System.Windows.Forms.Padding(4);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(136, 22);
            this.dateTimePicker1.TabIndex = 68;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker2.Location = new System.Drawing.Point(165, 15);
            this.dateTimePicker2.Margin = new System.Windows.Forms.Padding(4);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(136, 22);
            this.dateTimePicker2.TabIndex = 67;
            // 
            // cbo_department
            // 
            this.cbo_department.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_department.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_department.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_department.FormattingEnabled = true;
            this.cbo_department.Location = new System.Drawing.Point(153, 69);
            this.cbo_department.Margin = new System.Windows.Forms.Padding(4);
            this.cbo_department.Name = "cbo_department";
            this.cbo_department.Size = new System.Drawing.Size(455, 26);
            this.cbo_department.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 74);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(85, 18);
            this.label2.TabIndex = 5;
            this.label2.Text = "Department";
            // 
            // rpt_payroll_summary
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1604, 777);
            this.Controls.Add(this.tbcntrl_main);
            this.Controls.Add(this.pnl_sidebar);
            this.Controls.Add(this.pnl_rpt_option_header);
            this.Name = "rpt_payroll_summary";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Payroll Summary Report";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.rpt_payroll_summary_Load);
            this.pnl_sidebar.ResumeLayout(false);
            this.tbcntrl_option.ResumeLayout(false);
            this.tpg_opt_1.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.pnl_rpt_option_header.ResumeLayout(false);
            this.pnl_rpt_option.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pic_loading)).EndInit();
            this.grp_options.ResumeLayout(false);
            this.grp_options.PerformLayout();
            this.tbcntrl_main.ResumeLayout(false);
            this.tpg_list.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvl_payroll_summary)).EndInit();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnl_sidebar;
        private System.Windows.Forms.TabControl tbcntrl_option;
        private System.Windows.Forms.TabPage tpg_opt_1;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btn_deletefile;
        private System.Windows.Forms.Button btn_print;
        private System.Windows.Forms.DataGridViewTextBoxColumn empname;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_code;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_period;
        private System.ComponentModel.BackgroundWorker bgworker;
        private System.Windows.Forms.Panel pnl_rpt_option_header;
        private System.Windows.Forms.Panel pnl_rpt_option;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox pic_loading;
        private System.Windows.Forms.Button btn_submit;
        private System.Windows.Forms.GroupBox grp_options;
        private System.Windows.Forms.ComboBox cbo_payollperiod;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbo_employee;
        private System.Windows.Forms.Label lbl_cbo_1;
        private System.Windows.Forms.DataGridViewTextBoxColumn empid;
        private System.Windows.Forms.DataGridViewTextBoxColumn ppid;
        private System.Windows.Forms.TabControl tbcntrl_main;
        private System.Windows.Forms.TabPage tpg_list;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgvl_payroll_summary;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.DataGridViewTextBoxColumn rpt_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn filename;
        private System.Windows.Forms.DataGridViewTextBoxColumn date_added;
        private System.Windows.Forms.ComboBox cbo_department;
        private System.Windows.Forms.Label label2;
    }
}