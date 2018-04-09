namespace Human_Resource_Information_System
{
    partial class PaySlips
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tbcntrl_option = new System.Windows.Forms.TabControl();
            this.tpg_opt_1 = new System.Windows.Forms.TabPage();
            this.panel5 = new System.Windows.Forms.Panel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btn_deletefile = new System.Windows.Forms.Button();
            this.btn_print = new System.Windows.Forms.Button();
            this.empname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pay_code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgvl_payslips_files = new System.Windows.Forms.DataGridView();
            this.pay_slip_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.filename = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.date_created = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel7 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.tpg_list = new System.Windows.Forms.TabPage();
            this.tbcntrl_main = new System.Windows.Forms.TabControl();
            this.pnl_main = new System.Windows.Forms.Panel();
            this.empid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pnl_sidebar = new System.Windows.Forms.Panel();
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
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.ppid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgv_list = new System.Windows.Forms.DataGridView();
            this.tbcntrl_option.SuspendLayout();
            this.tpg_opt_1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvl_payslips_files)).BeginInit();
            this.panel7.SuspendLayout();
            this.panel4.SuspendLayout();
            this.tpg_list.SuspendLayout();
            this.tbcntrl_main.SuspendLayout();
            this.pnl_main.SuspendLayout();
            this.pnl_sidebar.SuspendLayout();
            this.pnl_rpt_option_header.SuspendLayout();
            this.pnl_rpt_option.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_loading)).BeginInit();
            this.grp_options.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).BeginInit();
            this.SuspendLayout();
            // 
            // tbcntrl_option
            // 
            this.tbcntrl_option.Alignment = System.Windows.Forms.TabAlignment.Bottom;
            this.tbcntrl_option.Controls.Add(this.tpg_opt_1);
            this.tbcntrl_option.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_option.Location = new System.Drawing.Point(0, 0);
            this.tbcntrl_option.Name = "tbcntrl_option";
            this.tbcntrl_option.SelectedIndex = 0;
            this.tbcntrl_option.Size = new System.Drawing.Size(197, 542);
            this.tbcntrl_option.TabIndex = 1;
            // 
            // tpg_opt_1
            // 
            this.tpg_opt_1.Controls.Add(this.panel5);
            this.tpg_opt_1.Location = new System.Drawing.Point(4, 4);
            this.tpg_opt_1.Name = "tpg_opt_1";
            this.tpg_opt_1.Size = new System.Drawing.Size(189, 516);
            this.tpg_opt_1.TabIndex = 2;
            this.tpg_opt_1.UseVisualStyleBackColor = true;
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.SteelBlue;
            this.panel5.Controls.Add(this.groupBox4);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(0, 0);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(189, 516);
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
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(189, 488);
            this.groupBox4.TabIndex = 64;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Main Option";
            // 
            // btn_deletefile
            // 
            this.btn_deletefile.BackColor = System.Drawing.Color.Maroon;
            this.btn_deletefile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_deletefile.ForeColor = System.Drawing.SystemColors.Info;
            this.btn_deletefile.Location = new System.Drawing.Point(8, 116);
            this.btn_deletefile.Name = "btn_deletefile";
            this.btn_deletefile.Size = new System.Drawing.Size(156, 72);
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
            this.btn_print.Location = new System.Drawing.Point(8, 38);
            this.btn_print.Name = "btn_print";
            this.btn_print.Size = new System.Drawing.Size(156, 72);
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
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgvl_payslips_files);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox2.Location = new System.Drawing.Point(0, 45);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(992, 465);
            this.groupBox2.TabIndex = 58;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Payroll Period List";
            // 
            // dgvl_payslips_files
            // 
            this.dgvl_payslips_files.AllowUserToResizeRows = false;
            this.dgvl_payslips_files.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvl_payslips_files.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.pay_slip_id,
            this.filename,
            this.date_created});
            this.dgvl_payslips_files.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvl_payslips_files.Location = new System.Drawing.Point(3, 18);
            this.dgvl_payslips_files.Name = "dgvl_payslips_files";
            this.dgvl_payslips_files.ReadOnly = true;
            this.dgvl_payslips_files.RowHeadersVisible = false;
            this.dgvl_payslips_files.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvl_payslips_files.Size = new System.Drawing.Size(986, 444);
            this.dgvl_payslips_files.TabIndex = 1;
            // 
            // pay_slip_id
            // 
            this.pay_slip_id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.pay_slip_id.HeaderText = "PAYSLIP CODE";
            this.pay_slip_id.Name = "pay_slip_id";
            this.pay_slip_id.ReadOnly = true;
            // 
            // filename
            // 
            this.filename.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.filename.HeaderText = "PAYSLIP FILENAME";
            this.filename.Name = "filename";
            this.filename.ReadOnly = true;
            // 
            // date_created
            // 
            this.date_created.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.date_created.HeaderText = "DATE CREATED";
            this.date_created.Name = "date_created";
            this.date_created.ReadOnly = true;
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.SystemColors.Info;
            this.panel7.Controls.Add(this.groupBox2);
            this.panel7.Controls.Add(this.panel4);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(3, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(992, 510);
            this.panel7.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.label5);
            this.panel4.Controls.Add(this.label19);
            this.panel4.Controls.Add(this.dateTimePicker1);
            this.panel4.Controls.Add(this.dateTimePicker2);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(992, 45);
            this.panel4.TabIndex = 59;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(233, 17);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(20, 13);
            this.label5.TabIndex = 70;
            this.label5.Text = "To";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(12, 17);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(94, 13);
            this.label19.TabIndex = 69;
            this.label19.Text = "Transaction Dates";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(264, 12);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(103, 20);
            this.dateTimePicker1.TabIndex = 68;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker2.Location = new System.Drawing.Point(124, 12);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(103, 20);
            this.dateTimePicker2.TabIndex = 67;
            // 
            // tpg_list
            // 
            this.tpg_list.Controls.Add(this.panel7);
            this.tpg_list.Location = new System.Drawing.Point(4, 22);
            this.tpg_list.Name = "tpg_list";
            this.tpg_list.Padding = new System.Windows.Forms.Padding(3);
            this.tpg_list.Size = new System.Drawing.Size(998, 516);
            this.tpg_list.TabIndex = 0;
            this.tpg_list.Text = "Generated Payroll List";
            this.tpg_list.UseVisualStyleBackColor = true;
            // 
            // tbcntrl_main
            // 
            this.tbcntrl_main.Controls.Add(this.tpg_list);
            this.tbcntrl_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_main.Location = new System.Drawing.Point(0, 0);
            this.tbcntrl_main.Name = "tbcntrl_main";
            this.tbcntrl_main.SelectedIndex = 0;
            this.tbcntrl_main.Size = new System.Drawing.Size(1006, 542);
            this.tbcntrl_main.TabIndex = 1;
            // 
            // pnl_main
            // 
            this.pnl_main.Controls.Add(this.tbcntrl_main);
            this.pnl_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnl_main.Location = new System.Drawing.Point(197, 89);
            this.pnl_main.Name = "pnl_main";
            this.pnl_main.Size = new System.Drawing.Size(1006, 542);
            this.pnl_main.TabIndex = 63;
            // 
            // empid
            // 
            this.empid.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.empid.HeaderText = "EMPLOYEE ID";
            this.empid.Name = "empid";
            this.empid.ReadOnly = true;
            // 
            // pnl_sidebar
            // 
            this.pnl_sidebar.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_sidebar.Controls.Add(this.tbcntrl_option);
            this.pnl_sidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnl_sidebar.Location = new System.Drawing.Point(0, 89);
            this.pnl_sidebar.Name = "pnl_sidebar";
            this.pnl_sidebar.Size = new System.Drawing.Size(197, 542);
            this.pnl_sidebar.TabIndex = 62;
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
            this.pnl_rpt_option_header.Name = "pnl_rpt_option_header";
            this.pnl_rpt_option_header.Size = new System.Drawing.Size(1203, 89);
            this.pnl_rpt_option_header.TabIndex = 61;
            // 
            // pnl_rpt_option
            // 
            this.pnl_rpt_option.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_rpt_option.Controls.Add(this.panel2);
            this.pnl_rpt_option.Controls.Add(this.grp_options);
            this.pnl_rpt_option.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_rpt_option.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnl_rpt_option.Location = new System.Drawing.Point(0, 0);
            this.pnl_rpt_option.Name = "pnl_rpt_option";
            this.pnl_rpt_option.Size = new System.Drawing.Size(1203, 86);
            this.pnl_rpt_option.TabIndex = 8;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.pic_loading);
            this.panel2.Controls.Add(this.btn_submit);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(630, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(171, 86);
            this.panel2.TabIndex = 4;
            // 
            // pic_loading
            // 
            this.pic_loading.ErrorImage = global::Human_Resource_Information_System.Properties.Resources.spin;
            this.pic_loading.Image = global::Human_Resource_Information_System.Properties.Resources.spin;
            this.pic_loading.Location = new System.Drawing.Point(61, 59);
            this.pic_loading.Name = "pic_loading";
            this.pic_loading.Size = new System.Drawing.Size(32, 26);
            this.pic_loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic_loading.TabIndex = 7;
            this.pic_loading.TabStop = false;

            // 
            // btn_submit
            // 
            this.btn_submit.BackColor = System.Drawing.Color.SaddleBrown;
            this.btn_submit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_submit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_submit.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btn_submit.Location = new System.Drawing.Point(6, 10);
            this.btn_submit.Name = "btn_submit";
            this.btn_submit.Size = new System.Drawing.Size(139, 43);
            this.btn_submit.TabIndex = 1;
            this.btn_submit.Text = "Generate Payslips";
            this.btn_submit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_submit.UseVisualStyleBackColor = false;
            this.btn_submit.Click += new System.EventHandler(this.btn_submit_Click);
            // 
            // grp_options
            // 
            this.grp_options.Controls.Add(this.cbo_payollperiod);
            this.grp_options.Controls.Add(this.label1);
            this.grp_options.Controls.Add(this.cbo_employee);
            this.grp_options.Controls.Add(this.lbl_cbo_1);
            this.grp_options.Dock = System.Windows.Forms.DockStyle.Left;
            this.grp_options.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grp_options.ForeColor = System.Drawing.SystemColors.Info;
            this.grp_options.Location = new System.Drawing.Point(0, 0);
            this.grp_options.Name = "grp_options";
            this.grp_options.Size = new System.Drawing.Size(630, 86);
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
            this.cbo_payollperiod.Location = new System.Drawing.Point(115, 20);
            this.cbo_payollperiod.Name = "cbo_payollperiod";
            this.cbo_payollperiod.Size = new System.Drawing.Size(342, 23);
            this.cbo_payollperiod.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 23);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(83, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Payroll Period";
            // 
            // cbo_employee
            // 
            this.cbo_employee.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_employee.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_employee.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_employee.FormattingEnabled = true;
            this.cbo_employee.Location = new System.Drawing.Point(115, 49);
            this.cbo_employee.Name = "cbo_employee";
            this.cbo_employee.Size = new System.Drawing.Size(342, 23);
            this.cbo_employee.TabIndex = 1;
            // 
            // lbl_cbo_1
            // 
            this.lbl_cbo_1.AutoSize = true;
            this.lbl_cbo_1.Location = new System.Drawing.Point(7, 53);
            this.lbl_cbo_1.Name = "lbl_cbo_1";
            this.lbl_cbo_1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_cbo_1.Size = new System.Drawing.Size(62, 15);
            this.lbl_cbo_1.TabIndex = 1;
            this.lbl_cbo_1.Text = "Employee";
            // 
            // groupBox6
            // 
            this.groupBox6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox6.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox6.Location = new System.Drawing.Point(0, 0);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(1203, 631);
            this.groupBox6.TabIndex = 64;
            this.groupBox6.TabStop = false;
            // 
            // ppid
            // 
            this.ppid.HeaderText = "ppid";
            this.ppid.Name = "ppid";
            this.ppid.ReadOnly = true;
            this.ppid.Visible = false;
            // 
            // dgv_list
            // 
            this.dgv_list.AllowUserToResizeRows = false;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_list.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.dgv_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_list.DefaultCellStyle = dataGridViewCellStyle8;
            this.dgv_list.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_list.Location = new System.Drawing.Point(0, 0);
            this.dgv_list.Name = "dgv_list";
            this.dgv_list.ReadOnly = true;
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_list.RowHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.dgv_list.RowHeadersVisible = false;
            this.dgv_list.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_list.Size = new System.Drawing.Size(1203, 631);
            this.dgv_list.TabIndex = 60;
            // 
            // PaySlips
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1203, 631);
            this.Controls.Add(this.pnl_main);
            this.Controls.Add(this.pnl_sidebar);
            this.Controls.Add(this.pnl_rpt_option_header);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.dgv_list);
            this.Name = "PaySlips";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PaySlips";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.PaySlips_Load);
            this.tbcntrl_option.ResumeLayout(false);
            this.tpg_opt_1.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvl_payslips_files)).EndInit();
            this.panel7.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.tpg_list.ResumeLayout(false);
            this.tbcntrl_main.ResumeLayout(false);
            this.pnl_main.ResumeLayout(false);
            this.pnl_sidebar.ResumeLayout(false);
            this.pnl_rpt_option_header.ResumeLayout(false);
            this.pnl_rpt_option.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pic_loading)).EndInit();
            this.grp_options.ResumeLayout(false);
            this.grp_options.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tbcntrl_option;
        private System.Windows.Forms.TabPage tpg_opt_1;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btn_print;
        private System.Windows.Forms.DataGridViewTextBoxColumn empname;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_code;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgvl_payslips_files;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.TabPage tpg_list;
        private System.Windows.Forms.TabControl tbcntrl_main;
        private System.Windows.Forms.Panel pnl_main;
        private System.Windows.Forms.DataGridViewTextBoxColumn empid;
        private System.Windows.Forms.Panel pnl_sidebar;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_period;
        private System.ComponentModel.BackgroundWorker bgworker;
        private System.Windows.Forms.Panel pnl_rpt_option_header;
        private System.Windows.Forms.Panel pnl_rpt_option;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btn_submit;
        private System.Windows.Forms.GroupBox grp_options;
        private System.Windows.Forms.ComboBox cbo_payollperiod;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbo_employee;
        private System.Windows.Forms.Label lbl_cbo_1;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.DataGridViewTextBoxColumn ppid;
        private System.Windows.Forms.DataGridView dgv_list;
        private System.Windows.Forms.PictureBox pic_loading;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_slip_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn filename;
        private System.Windows.Forms.DataGridViewTextBoxColumn date_created;
        private System.Windows.Forms.Button btn_deletefile;


    }
}