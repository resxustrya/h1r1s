namespace Human_Resource_Information_System
{
    partial class m_ShiftSchedule
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
            this.label4 = new System.Windows.Forms.Label();
            this.txt_code = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.tpg_info = new System.Windows.Forms.TabPage();
            this.cbo_type = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_remark = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chk_am = new System.Windows.Forms.CheckBox();
            this.chk_pm = new System.Windows.Forms.CheckBox();
            this.dtp_pm_out = new System.Windows.Forms.DateTimePicker();
            this.dtp_pm_in = new System.Windows.Forms.DateTimePicker();
            this.dtp_am_out = new System.Windows.Forms.DateTimePicker();
            this.dtp_am_in = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.tpg_list = new System.Windows.Forms.TabPage();
            this.dgv_list = new System.Windows.Forms.DataGridView();
            this.tbcntrl_main = new System.Windows.Forms.TabControl();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btn_save = new System.Windows.Forms.Button();
            this.btn_back = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btn_print = new System.Windows.Forms.Button();
            this.btn_additem = new System.Windows.Forms.Button();
            this.btn_delitem = new System.Windows.Forms.Button();
            this.btn_upditem = new System.Windows.Forms.Button();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.cbo_searchby = new System.Windows.Forms.ComboBox();
            this.txt_search = new System.Windows.Forms.TextBox();
            this.btn_search = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.tpg_opt_1 = new System.Windows.Forms.TabPage();
            this.tbcntrl_option = new System.Windows.Forms.TabControl();
            this.tpg_opt_2 = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timein_am = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timeout_am = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timein_pm = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.timeout_pm = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.remarks = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tpg_info.SuspendLayout();
            this.tpg_list.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).BeginInit();
            this.tbcntrl_main.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tpg_opt_1.SuspendLayout();
            this.tbcntrl_option.SuspendLayout();
            this.tpg_opt_2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(263, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(16, 20);
            this.label4.TabIndex = 46;
            this.label4.Text = "*";
            // 
            // txt_code
            // 
            this.txt_code.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt_code.Location = new System.Drawing.Point(133, 9);
            this.txt_code.MaxLength = 8;
            this.txt_code.Name = "txt_code";
            this.txt_code.Size = new System.Drawing.Size(123, 20);
            this.txt_code.TabIndex = 35;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 40);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(115, 13);
            this.label10.TabIndex = 33;
            this.label10.Text = "Time In/Time Out (AM)";
            // 
            // tpg_info
            // 
            this.tpg_info.BackColor = System.Drawing.SystemColors.Info;
            this.tpg_info.Controls.Add(this.cbo_type);
            this.tpg_info.Controls.Add(this.label3);
            this.tpg_info.Controls.Add(this.label2);
            this.tpg_info.Controls.Add(this.txt_remark);
            this.tpg_info.Controls.Add(this.label1);
            this.tpg_info.Controls.Add(this.chk_am);
            this.tpg_info.Controls.Add(this.chk_pm);
            this.tpg_info.Controls.Add(this.dtp_pm_out);
            this.tpg_info.Controls.Add(this.dtp_pm_in);
            this.tpg_info.Controls.Add(this.dtp_am_out);
            this.tpg_info.Controls.Add(this.dtp_am_in);
            this.tpg_info.Controls.Add(this.label4);
            this.tpg_info.Controls.Add(this.txt_code);
            this.tpg_info.Controls.Add(this.label8);
            this.tpg_info.Controls.Add(this.label10);
            this.tpg_info.Location = new System.Drawing.Point(4, 22);
            this.tpg_info.Name = "tpg_info";
            this.tpg_info.Padding = new System.Windows.Forms.Padding(3);
            this.tpg_info.Size = new System.Drawing.Size(536, 481);
            this.tpg_info.TabIndex = 2;
            this.tpg_info.Text = "Shift Schedule Info";
            // 
            // cbo_type
            // 
            this.cbo_type.BackColor = System.Drawing.Color.White;
            this.cbo_type.FormattingEnabled = true;
            this.cbo_type.Items.AddRange(new object[] {
            "Regular",
            "Overtime"});
            this.cbo_type.Location = new System.Drawing.Point(133, 109);
            this.cbo_type.Name = "cbo_type";
            this.cbo_type.Size = new System.Drawing.Size(192, 21);
            this.cbo_type.TabIndex = 104;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 112);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 103;
            this.label3.Text = "Type";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 102;
            this.label2.Text = "Remarks";
            // 
            // txt_remark
            // 
            this.txt_remark.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt_remark.Location = new System.Drawing.Point(133, 85);
            this.txt_remark.MaxLength = 8;
            this.txt_remark.Name = "txt_remark";
            this.txt_remark.Size = new System.Drawing.Size(232, 20);
            this.txt_remark.TabIndex = 101;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 13);
            this.label1.TabIndex = 100;
            this.label1.Text = "Time In/Time Out (PM)";
            // 
            // chk_am
            // 
            this.chk_am.AutoSize = true;
            this.chk_am.Checked = true;
            this.chk_am.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk_am.Location = new System.Drawing.Point(133, 40);
            this.chk_am.Name = "chk_am";
            this.chk_am.Size = new System.Drawing.Size(15, 14);
            this.chk_am.TabIndex = 99;
            this.chk_am.UseVisualStyleBackColor = true;
            this.chk_am.CheckedChanged += new System.EventHandler(this.chk_am_CheckedChanged);
            // 
            // chk_pm
            // 
            this.chk_pm.AutoSize = true;
            this.chk_pm.Location = new System.Drawing.Point(133, 65);
            this.chk_pm.Name = "chk_pm";
            this.chk_pm.Size = new System.Drawing.Size(15, 14);
            this.chk_pm.TabIndex = 98;
            this.chk_pm.UseVisualStyleBackColor = true;
            this.chk_pm.CheckedChanged += new System.EventHandler(this.chk_pm_CheckedChanged);
            // 
            // dtp_pm_out
            // 
            this.dtp_pm_out.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtp_pm_out.Location = new System.Drawing.Point(263, 61);
            this.dtp_pm_out.Name = "dtp_pm_out";
            this.dtp_pm_out.Size = new System.Drawing.Size(103, 20);
            this.dtp_pm_out.TabIndex = 97;
            // 
            // dtp_pm_in
            // 
            this.dtp_pm_in.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtp_pm_in.Location = new System.Drawing.Point(154, 61);
            this.dtp_pm_in.Name = "dtp_pm_in";
            this.dtp_pm_in.Size = new System.Drawing.Size(103, 20);
            this.dtp_pm_in.TabIndex = 95;
            // 
            // dtp_am_out
            // 
            this.dtp_am_out.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtp_am_out.Location = new System.Drawing.Point(263, 35);
            this.dtp_am_out.Name = "dtp_am_out";
            this.dtp_am_out.Size = new System.Drawing.Size(103, 20);
            this.dtp_am_out.TabIndex = 93;
            // 
            // dtp_am_in
            // 
            this.dtp_am_in.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtp_am_in.Location = new System.Drawing.Point(154, 35);
            this.dtp_am_in.Name = "dtp_am_in";
            this.dtp_am_in.Size = new System.Drawing.Size(103, 20);
            this.dtp_am_in.TabIndex = 92;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(10, 13);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(32, 13);
            this.label8.TabIndex = 32;
            this.label8.Text = "Code";
            // 
            // tpg_list
            // 
            this.tpg_list.Controls.Add(this.dgv_list);
            this.tpg_list.Location = new System.Drawing.Point(4, 22);
            this.tpg_list.Name = "tpg_list";
            this.tpg_list.Padding = new System.Windows.Forms.Padding(3);
            this.tpg_list.Size = new System.Drawing.Size(536, 481);
            this.tpg_list.TabIndex = 1;
            this.tpg_list.Text = "Shift Schedule List";
            this.tpg_list.UseVisualStyleBackColor = true;
            // 
            // dgv_list
            // 
            this.dgv_list.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_list.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.code,
            this.timein_am,
            this.timeout_am,
            this.timein_pm,
            this.timeout_pm,
            this.remarks,
            this.type});
            this.dgv_list.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_list.Location = new System.Drawing.Point(3, 3);
            this.dgv_list.Name = "dgv_list";
            this.dgv_list.ReadOnly = true;
            this.dgv_list.RowHeadersWidth = 25;
            this.dgv_list.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_list.Size = new System.Drawing.Size(530, 475);
            this.dgv_list.TabIndex = 1;
            // 
            // tbcntrl_main
            // 
            this.tbcntrl_main.Controls.Add(this.tpg_list);
            this.tbcntrl_main.Controls.Add(this.tpg_info);
            this.tbcntrl_main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_main.Location = new System.Drawing.Point(0, 0);
            this.tbcntrl_main.Name = "tbcntrl_main";
            this.tbcntrl_main.SelectedIndex = 0;
            this.tbcntrl_main.Size = new System.Drawing.Size(544, 507);
            this.tbcntrl_main.TabIndex = 7;
            this.tbcntrl_main.Selecting += new System.Windows.Forms.TabControlCancelEventHandler(this.tbcntrl_main_Selecting);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tbcntrl_main);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(221, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(544, 507);
            this.panel2.TabIndex = 9;
            // 
            // btn_save
            // 
            this.btn_save.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_save.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_save.Location = new System.Drawing.Point(12, 24);
            this.btn_save.Name = "btn_save";
            this.btn_save.Size = new System.Drawing.Size(160, 46);
            this.btn_save.TabIndex = 8;
            this.btn_save.Text = "Save";
            this.btn_save.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_save.UseVisualStyleBackColor = false;
            this.btn_save.Click += new System.EventHandler(this.btn_save_Click);
            // 
            // btn_back
            // 
            this.btn_back.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_back.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_back.Location = new System.Drawing.Point(12, 73);
            this.btn_back.Name = "btn_back";
            this.btn_back.Size = new System.Drawing.Size(160, 46);
            this.btn_back.TabIndex = 1;
            this.btn_back.Text = "Back";
            this.btn_back.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_back.UseVisualStyleBackColor = false;
            this.btn_back.Click += new System.EventHandler(this.btn_back_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btn_save);
            this.groupBox3.Controls.Add(this.btn_back);
            this.groupBox3.ForeColor = System.Drawing.SystemColors.Info;
            this.groupBox3.Location = new System.Drawing.Point(10, 14);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(184, 130);
            this.groupBox3.TabIndex = 48;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Option";
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.SteelBlue;
            this.panel4.Controls.Add(this.groupBox3);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(3, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(207, 475);
            this.panel4.TabIndex = 0;
            // 
            // btn_print
            // 
            this.btn_print.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_print.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_print.Location = new System.Drawing.Point(12, 204);
            this.btn_print.Name = "btn_print";
            this.btn_print.Size = new System.Drawing.Size(160, 55);
            this.btn_print.TabIndex = 9;
            this.btn_print.Text = "Print List";
            this.btn_print.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_print.UseVisualStyleBackColor = false;
            this.btn_print.Click += new System.EventHandler(this.btn_print_Click);
            // 
            // btn_additem
            // 
            this.btn_additem.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_additem.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_additem.Location = new System.Drawing.Point(12, 21);
            this.btn_additem.Name = "btn_additem";
            this.btn_additem.Size = new System.Drawing.Size(160, 55);
            this.btn_additem.TabIndex = 8;
            this.btn_additem.Text = "Add New";
            this.btn_additem.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_additem.UseVisualStyleBackColor = false;
            this.btn_additem.Click += new System.EventHandler(this.btn_additem_Click);
            // 
            // btn_delitem
            // 
            this.btn_delitem.BackColor = System.Drawing.Color.Maroon;
            this.btn_delitem.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_delitem.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btn_delitem.Location = new System.Drawing.Point(12, 143);
            this.btn_delitem.Name = "btn_delitem";
            this.btn_delitem.Size = new System.Drawing.Size(160, 55);
            this.btn_delitem.TabIndex = 7;
            this.btn_delitem.Text = "Delete";
            this.btn_delitem.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_delitem.UseVisualStyleBackColor = false;
            this.btn_delitem.Click += new System.EventHandler(this.btn_delitem_Click);
            // 
            // btn_upditem
            // 
            this.btn_upditem.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_upditem.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_upditem.Location = new System.Drawing.Point(12, 82);
            this.btn_upditem.Name = "btn_upditem";
            this.btn_upditem.Size = new System.Drawing.Size(160, 55);
            this.btn_upditem.TabIndex = 1;
            this.btn_upditem.Text = "Update";
            this.btn_upditem.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_upditem.UseVisualStyleBackColor = false;
            this.btn_upditem.Click += new System.EventHandler(this.btn_upditem_Click);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.btn_print);
            this.groupBox7.Controls.Add(this.btn_additem);
            this.groupBox7.Controls.Add(this.btn_delitem);
            this.groupBox7.Controls.Add(this.btn_upditem);
            this.groupBox7.ForeColor = System.Drawing.SystemColors.Info;
            this.groupBox7.Location = new System.Drawing.Point(11, 13);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(184, 273);
            this.groupBox7.TabIndex = 46;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Main Option";
            // 
            // cbo_searchby
            // 
            this.cbo_searchby.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_searchby.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_searchby.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbo_searchby.FormattingEnabled = true;
            this.cbo_searchby.Location = new System.Drawing.Point(12, 33);
            this.cbo_searchby.Name = "cbo_searchby";
            this.cbo_searchby.Size = new System.Drawing.Size(160, 21);
            this.cbo_searchby.TabIndex = 10;
            // 
            // txt_search
            // 
            this.txt_search.Location = new System.Drawing.Point(12, 72);
            this.txt_search.Name = "txt_search";
            this.txt_search.Size = new System.Drawing.Size(160, 20);
            this.txt_search.TabIndex = 9;
            // 
            // btn_search
            // 
            this.btn_search.BackColor = System.Drawing.Color.DarkSlateGray;
            this.btn_search.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_search.Location = new System.Drawing.Point(12, 100);
            this.btn_search.Name = "btn_search";
            this.btn_search.Size = new System.Drawing.Size(160, 55);
            this.btn_search.TabIndex = 8;
            this.btn_search.Text = "Search Now";
            this.btn_search.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_search.UseVisualStyleBackColor = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbo_searchby);
            this.groupBox2.Controls.Add(this.txt_search);
            this.groupBox2.Controls.Add(this.btn_search);
            this.groupBox2.ForeColor = System.Drawing.SystemColors.Info;
            this.groupBox2.Location = new System.Drawing.Point(11, 292);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(184, 165);
            this.groupBox2.TabIndex = 47;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Search Option";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.SteelBlue;
            this.panel3.Controls.Add(this.groupBox2);
            this.panel3.Controls.Add(this.groupBox7);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(3, 3);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(207, 475);
            this.panel3.TabIndex = 0;
            // 
            // tpg_opt_1
            // 
            this.tpg_opt_1.Controls.Add(this.panel3);
            this.tpg_opt_1.Location = new System.Drawing.Point(4, 4);
            this.tpg_opt_1.Name = "tpg_opt_1";
            this.tpg_opt_1.Padding = new System.Windows.Forms.Padding(3);
            this.tpg_opt_1.Size = new System.Drawing.Size(213, 481);
            this.tpg_opt_1.TabIndex = 0;
            this.tpg_opt_1.UseVisualStyleBackColor = true;
            // 
            // tbcntrl_option
            // 
            this.tbcntrl_option.Alignment = System.Windows.Forms.TabAlignment.Bottom;
            this.tbcntrl_option.Controls.Add(this.tpg_opt_1);
            this.tbcntrl_option.Controls.Add(this.tpg_opt_2);
            this.tbcntrl_option.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcntrl_option.Location = new System.Drawing.Point(0, 0);
            this.tbcntrl_option.Name = "tbcntrl_option";
            this.tbcntrl_option.SelectedIndex = 0;
            this.tbcntrl_option.Size = new System.Drawing.Size(221, 507);
            this.tbcntrl_option.TabIndex = 49;
            this.tbcntrl_option.Selecting += new System.Windows.Forms.TabControlCancelEventHandler(this.tbcntrl_option_Selecting);
            // 
            // tpg_opt_2
            // 
            this.tpg_opt_2.BackColor = System.Drawing.SystemColors.HotTrack;
            this.tpg_opt_2.Controls.Add(this.panel4);
            this.tpg_opt_2.Location = new System.Drawing.Point(4, 4);
            this.tpg_opt_2.Name = "tpg_opt_2";
            this.tpg_opt_2.Padding = new System.Windows.Forms.Padding(3);
            this.tpg_opt_2.Size = new System.Drawing.Size(213, 481);
            this.tpg_opt_2.TabIndex = 1;
            this.tpg_opt_2.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.SteelBlue;
            this.panel1.Controls.Add(this.tbcntrl_option);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(221, 507);
            this.panel1.TabIndex = 8;
            // 
            // code
            // 
            this.code.HeaderText = "CODE";
            this.code.Name = "code";
            this.code.ReadOnly = true;
            // 
            // timein_am
            // 
            this.timein_am.HeaderText = "TIME IN(AM)";
            this.timein_am.Name = "timein_am";
            this.timein_am.ReadOnly = true;
            // 
            // timeout_am
            // 
            this.timeout_am.HeaderText = "TIME OUT(AM)";
            this.timeout_am.Name = "timeout_am";
            this.timeout_am.ReadOnly = true;
            this.timeout_am.Width = 120;
            // 
            // timein_pm
            // 
            this.timein_pm.HeaderText = "TIME IN(PM)";
            this.timein_pm.Name = "timein_pm";
            this.timein_pm.ReadOnly = true;
            // 
            // timeout_pm
            // 
            this.timeout_pm.HeaderText = "TIME OUT(PM)";
            this.timeout_pm.Name = "timeout_pm";
            this.timeout_pm.ReadOnly = true;
            this.timeout_pm.Width = 120;
            // 
            // remarks
            // 
            this.remarks.HeaderText = "REMARKS";
            this.remarks.Name = "remarks";
            this.remarks.ReadOnly = true;
            // 
            // type
            // 
            this.type.HeaderText = "TYPE";
            this.type.Name = "type";
            this.type.ReadOnly = true;
            // 
            // m_ShiftSchedule
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(765, 507);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "m_ShiftSchedule";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Shift Schedule";
            this.Load += new System.EventHandler(this.m_ShiftSchedule_Load);
            this.tpg_info.ResumeLayout(false);
            this.tpg_info.PerformLayout();
            this.tpg_list.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_list)).EndInit();
            this.tbcntrl_main.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.tpg_opt_1.ResumeLayout(false);
            this.tbcntrl_option.ResumeLayout(false);
            this.tpg_opt_2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_code;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TabPage tpg_info;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TabPage tpg_list;
        private System.Windows.Forms.DataGridView dgv_list;
        private System.Windows.Forms.TabControl tbcntrl_main;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btn_save;
        private System.Windows.Forms.Button btn_back;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button btn_print;
        private System.Windows.Forms.Button btn_additem;
        private System.Windows.Forms.Button btn_delitem;
        private System.Windows.Forms.Button btn_upditem;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.ComboBox cbo_searchby;
        private System.Windows.Forms.TextBox txt_search;
        private System.Windows.Forms.Button btn_search;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TabPage tpg_opt_1;
        private System.Windows.Forms.TabControl tbcntrl_option;
        private System.Windows.Forms.TabPage tpg_opt_2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DateTimePicker dtp_pm_out;
        private System.Windows.Forms.DateTimePicker dtp_pm_in;
        private System.Windows.Forms.DateTimePicker dtp_am_out;
        private System.Windows.Forms.DateTimePicker dtp_am_in;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chk_am;
        private System.Windows.Forms.CheckBox chk_pm;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_remark;
        private System.Windows.Forms.ComboBox cbo_type;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridViewTextBoxColumn code;
        private System.Windows.Forms.DataGridViewTextBoxColumn timein_am;
        private System.Windows.Forms.DataGridViewTextBoxColumn timeout_am;
        private System.Windows.Forms.DataGridViewTextBoxColumn timein_pm;
        private System.Windows.Forms.DataGridViewTextBoxColumn timeout_pm;
        private System.Windows.Forms.DataGridViewTextBoxColumn remarks;
        private System.Windows.Forms.DataGridViewTextBoxColumn type;
    }
}