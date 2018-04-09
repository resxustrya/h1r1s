namespace Human_Resource_Information_System
{
    partial class rpt_print_dtr
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
            this.panel3 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.txt_search = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dtp_to = new System.Windows.Forms.DateTimePicker();
            this.dtp_frm = new System.Windows.Forms.DateTimePicker();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.dgv_list = new System.Windows.Forms.DataGridView();
            this.ppid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pay_period = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.empid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.empname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pay_code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tpg_opt_1 = new System.Windows.Forms.TabPage();
            this.panel5 = new System.Windows.Forms.Panel();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btn_deletefile = new System.Windows.Forms.Button();
            this.btn_print = new System.Windows.Forms.Button();
            this.dgvl_dtrfiles = new System.Windows.Forms.DataGridView();
            this.dtr_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.filename = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.date_created = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panel7 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.tpg_list = new System.Windows.Forms.TabPage();
            this.tbcntrl_main = new System.Windows.Forms.TabControl();
            this.pnl_main = new System.Windows.Forms.Panel();
            this.tbcntrl_option = new System.Windows.Forms.TabControl();
            this.pnl_sidebar = new System.Windows.Forms.Panel();
            this.pnl_rpt_option_header.SuspendLayout();
            this.pnl_rpt_option.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pic_loading)).BeginInit();
            this.grp_options.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).BeginInit();
            this.tpg_opt_1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvl_dtrfiles)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel4.SuspendLayout();
            this.tpg_list.SuspendLayout();
            this.tbcntrl_main.SuspendLayout();
            this.pnl_main.SuspendLayout();
            this.tbcntrl_option.SuspendLayout();
            this.pnl_sidebar.SuspendLayout();
            this.SuspendLayout();
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
            this.pnl_rpt_option_header.Size = new System.Drawing.Size(1203, 114);
            this.pnl_rpt_option_header.TabIndex = 15;
            // 
            // pnl_rpt_option
            // 
            this.pnl_rpt_option.BackColor = System.Drawing.Color.DarkKhaki;
            this.pnl_rpt_option.Controls.Add(this.panel2);
            this.pnl_rpt_option.Controls.Add(this.grp_options);
            this.pnl_rpt_option.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_rpt_option.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnl_rpt_option.Location = new System.Drawing.Point(0, 0);
            this.pnl_rpt_option.Name = "pnl_rpt_option";
            this.pnl_rpt_option.Size = new System.Drawing.Size(1203, 112);
            this.pnl_rpt_option.TabIndex = 8;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.pic_loading);
            this.panel2.Controls.Add(this.btn_submit);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(630, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(148, 112);
            this.panel2.TabIndex = 4;
            // 
            // pic_loading
            // 
            this.pic_loading.ErrorImage = global::Human_Resource_Information_System.Properties.Resources.spin;
            this.pic_loading.Image = global::Human_Resource_Information_System.Properties.Resources.spin;
            this.pic_loading.Location = new System.Drawing.Point(58, 60);
            this.pic_loading.Name = "pic_loading";
            this.pic_loading.Size = new System.Drawing.Size(32, 26);
            this.pic_loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pic_loading.TabIndex = 6;
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
            this.btn_submit.Text = "Generate DTR";
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
            this.grp_options.Size = new System.Drawing.Size(630, 112);
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
            this.cbo_payollperiod.Location = new System.Drawing.Point(119, 30);
            this.cbo_payollperiod.Name = "cbo_payollperiod";
            this.cbo_payollperiod.Size = new System.Drawing.Size(342, 23);
            this.cbo_payollperiod.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 33);
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
            this.cbo_employee.Location = new System.Drawing.Point(119, 63);
            this.cbo_employee.Name = "cbo_employee";
            this.cbo_employee.Size = new System.Drawing.Size(342, 23);
            this.cbo_employee.TabIndex = 1;
            // 
            // lbl_cbo_1
            // 
            this.lbl_cbo_1.AutoSize = true;
            this.lbl_cbo_1.Location = new System.Drawing.Point(11, 68);
            this.lbl_cbo_1.Name = "lbl_cbo_1";
            this.lbl_cbo_1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_cbo_1.Size = new System.Drawing.Size(62, 15);
            this.lbl_cbo_1.TabIndex = 1;
            this.lbl_cbo_1.Text = "Employee";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.button1);
            this.panel3.Controls.Add(this.txt_search);
            this.panel3.Controls.Add(this.label3);
            this.panel3.Controls.Add(this.label4);
            this.panel3.Controls.Add(this.dtp_to);
            this.panel3.Controls.Add(this.dtp_frm);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(992, 45);
            this.panel3.TabIndex = 59;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.OrangeRed;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.SystemColors.Info;
            this.button1.Location = new System.Drawing.Point(904, 10);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(73, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Search";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = false;
            // 
            // txt_search
            // 
            this.txt_search.Location = new System.Drawing.Point(719, 11);
            this.txt_search.Name = "txt_search";
            this.txt_search.Size = new System.Drawing.Size(179, 20);
            this.txt_search.TabIndex = 71;
            this.txt_search.Tag = "Search Employee";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(233, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 70;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 17);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 69;
            // 
            // dtp_to
            // 
            this.dtp_to.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtp_to.Location = new System.Drawing.Point(264, 12);
            this.dtp_to.Name = "dtp_to";
            this.dtp_to.Size = new System.Drawing.Size(103, 20);
            this.dtp_to.TabIndex = 68;
            // 
            // dtp_frm
            // 
            this.dtp_frm.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtp_frm.Location = new System.Drawing.Point(124, 12);
            this.dtp_frm.Name = "dtp_frm";
            this.dtp_frm.Size = new System.Drawing.Size(103, 20);
            this.dtp_frm.TabIndex = 67;
            // 
            // groupBox6
            // 
            this.groupBox6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox6.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox6.Location = new System.Drawing.Point(0, 45);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(992, 440);
            this.groupBox6.TabIndex = 58;
            this.groupBox6.TabStop = false;
            // 
            // dgv_list
            // 
            this.dgv_list.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_list.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_list.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv_list.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_list.Location = new System.Drawing.Point(3, 18);
            this.dgv_list.Name = "dgv_list";
            this.dgv_list.ReadOnly = true;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_list.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv_list.RowHeadersVisible = false;
            this.dgv_list.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_list.Size = new System.Drawing.Size(986, 419);
            this.dgv_list.TabIndex = 1;
            // 
            // ppid
            // 
            this.ppid.HeaderText = "ppid";
            this.ppid.Name = "ppid";
            this.ppid.ReadOnly = true;
            this.ppid.Visible = false;
            // 
            // pay_period
            // 
            this.pay_period.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.pay_period.HeaderText = "PAYROLL PERIOD";
            this.pay_period.Name = "pay_period";
            this.pay_period.ReadOnly = true;
            // 
            // empid
            // 
            this.empid.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.empid.HeaderText = "EMPLOYEE ID";
            this.empid.Name = "empid";
            this.empid.ReadOnly = true;
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
            // tpg_opt_1
            // 
            this.tpg_opt_1.Controls.Add(this.panel5);
            this.tpg_opt_1.Location = new System.Drawing.Point(4, 4);
            this.tpg_opt_1.Name = "tpg_opt_1";
            this.tpg_opt_1.Size = new System.Drawing.Size(189, 491);
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
            this.panel5.Size = new System.Drawing.Size(189, 491);
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
            this.btn_deletefile.Location = new System.Drawing.Point(8, 129);
            this.btn_deletefile.Name = "btn_deletefile";
            this.btn_deletefile.Size = new System.Drawing.Size(156, 72);
            this.btn_deletefile.TabIndex = 7;
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
            // dgvl_dtrfiles
            // 
            this.dgvl_dtrfiles.AllowUserToResizeRows = false;
            this.dgvl_dtrfiles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvl_dtrfiles.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dtr_id,
            this.filename,
            this.date_created});
            this.dgvl_dtrfiles.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvl_dtrfiles.Location = new System.Drawing.Point(3, 18);
            this.dgvl_dtrfiles.Name = "dgvl_dtrfiles";
            this.dgvl_dtrfiles.ReadOnly = true;
            this.dgvl_dtrfiles.RowHeadersVisible = false;
            this.dgvl_dtrfiles.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvl_dtrfiles.Size = new System.Drawing.Size(986, 419);
            this.dgvl_dtrfiles.TabIndex = 1;
            // 
            // dtr_id
            // 
            this.dtr_id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dtr_id.HeaderText = "DTR CODE";
            this.dtr_id.Name = "dtr_id";
            this.dtr_id.ReadOnly = true;
            // 
            // filename
            // 
            this.filename.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.filename.HeaderText = "DTR FILENAME";
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
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgvl_dtrfiles);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.groupBox2.Location = new System.Drawing.Point(0, 45);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(992, 440);
            this.groupBox2.TabIndex = 58;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Payroll Period List";
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.SystemColors.Info;
            this.panel7.Controls.Add(this.groupBox2);
            this.panel7.Controls.Add(this.panel4);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(3, 3);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(992, 485);
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
            this.tpg_list.Size = new System.Drawing.Size(998, 491);
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
            this.tbcntrl_main.Size = new System.Drawing.Size(1006, 517);
            this.tbcntrl_main.TabIndex = 1;
            // 
            // pnl_main
            // 
            this.pnl_main.Controls.Add(this.tbcntrl_main);
            this.pnl_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnl_main.Location = new System.Drawing.Point(197, 114);
            this.pnl_main.Name = "pnl_main";
            this.pnl_main.Size = new System.Drawing.Size(1006, 517);
            this.pnl_main.TabIndex = 52;
            // 
            // tbcntrl_option
            // 
            this.tbcntrl_option.Alignment = System.Windows.Forms.TabAlignment.Bottom;
            this.tbcntrl_option.Controls.Add(this.tpg_opt_1);
            this.tbcntrl_option.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_option.Location = new System.Drawing.Point(0, 0);
            this.tbcntrl_option.Name = "tbcntrl_option";
            this.tbcntrl_option.SelectedIndex = 0;
            this.tbcntrl_option.Size = new System.Drawing.Size(197, 517);
            this.tbcntrl_option.TabIndex = 1;
            // 
            // pnl_sidebar
            // 
            this.pnl_sidebar.BackColor = System.Drawing.Color.SteelBlue;
            this.pnl_sidebar.Controls.Add(this.tbcntrl_option);
            this.pnl_sidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnl_sidebar.Location = new System.Drawing.Point(0, 114);
            this.pnl_sidebar.Name = "pnl_sidebar";
            this.pnl_sidebar.Size = new System.Drawing.Size(197, 517);
            this.pnl_sidebar.TabIndex = 51;
            // 
            // rpt_print_dtr
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1203, 631);
            this.Controls.Add(this.pnl_main);
            this.Controls.Add(this.pnl_sidebar);
            this.Controls.Add(this.pnl_rpt_option_header);
            this.Name = "rpt_print_dtr";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Print Daily Time Record Report";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.rpt_print_dtr_Load);
            this.pnl_rpt_option_header.ResumeLayout(false);
            this.pnl_rpt_option.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pic_loading)).EndInit();
            this.grp_options.ResumeLayout(false);
            this.grp_options.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).EndInit();
            this.tpg_opt_1.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvl_dtrfiles)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.tpg_list.ResumeLayout(false);
            this.tbcntrl_main.ResumeLayout(false);
            this.pnl_main.ResumeLayout(false);
            this.tbcntrl_option.ResumeLayout(false);
            this.pnl_sidebar.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

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
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txt_search;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dtp_to;
        private System.Windows.Forms.DateTimePicker dtp_frm;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.DataGridView dgv_list;
        private System.Windows.Forms.DataGridViewTextBoxColumn ppid;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_period;
        private System.Windows.Forms.DataGridViewTextBoxColumn empid;
        private System.Windows.Forms.DataGridViewTextBoxColumn empname;
        private System.Windows.Forms.DataGridViewTextBoxColumn pay_code;
        private System.Windows.Forms.TabPage tpg_opt_1;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button btn_print;
        private System.Windows.Forms.DataGridView dgvl_dtrfiles;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.TabPage tpg_list;
        private System.Windows.Forms.TabControl tbcntrl_main;
        private System.Windows.Forms.Panel pnl_main;
        private System.Windows.Forms.TabControl tbcntrl_option;
        private System.Windows.Forms.Panel pnl_sidebar;
        private System.Windows.Forms.DataGridViewTextBoxColumn dtr_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn filename;
        private System.Windows.Forms.DataGridViewTextBoxColumn date_created;
        private System.Windows.Forms.Button btn_deletefile;
        private System.Windows.Forms.PictureBox pic_loading;
    }
}