namespace Human_Resource_Information_System
{
    partial class RPT_RES_entry
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
            this.bgworker = new System.ComponentModel.BackgroundWorker();
            this.pnl_rpt_option_header = new System.Windows.Forms.Panel();
            this.pnl_rpt_option = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btn_clear = new System.Windows.Forms.Button();
            this.grp_options = new System.Windows.Forms.GroupBox();
            this.cbo_3 = new System.Windows.Forms.ComboBox();
            this.lbl_cbo_3 = new System.Windows.Forms.Label();
            this.cbo_2 = new System.Windows.Forms.ComboBox();
            this.lbl_cbo_2 = new System.Windows.Forms.Label();
            this.chk_3 = new System.Windows.Forms.CheckBox();
            this.chk_2 = new System.Windows.Forms.CheckBox();
            this.chk_1 = new System.Windows.Forms.CheckBox();
            this.cbo_1 = new System.Windows.Forms.ComboBox();
            this.lbl_cbo_1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.grp_bybranch = new System.Windows.Forms.GroupBox();
            this.cbo_4 = new System.Windows.Forms.ComboBox();
            this.grp_bydate = new System.Windows.Forms.GroupBox();
            this.lbl_dt_to = new System.Windows.Forms.Label();
            this.dtp_to = new System.Windows.Forms.DateTimePicker();
            this.lbl_dt_frm = new System.Windows.Forms.Label();
            this.dtp_frm = new System.Windows.Forms.DateTimePicker();
            this.btn_minimize = new System.Windows.Forms.Button();
            this.pnl_pbar = new System.Windows.Forms.Panel();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.crptviewer = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.btn_submit = new System.Windows.Forms.Button();
            this.pnl_rpt_option_header.SuspendLayout();
            this.pnl_rpt_option.SuspendLayout();
            this.panel2.SuspendLayout();
            this.grp_options.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grp_bybranch.SuspendLayout();
            this.grp_bydate.SuspendLayout();
            this.pnl_pbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // bgworker
            // 
            this.bgworker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgworker_DoWork);
            // 
            // pnl_rpt_option_header
            // 
            this.pnl_rpt_option_header.Controls.Add(this.pnl_rpt_option);
            this.pnl_rpt_option_header.Controls.Add(this.btn_minimize);
            this.pnl_rpt_option_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_rpt_option_header.Location = new System.Drawing.Point(0, 0);
            this.pnl_rpt_option_header.Name = "pnl_rpt_option_header";
            this.pnl_rpt_option_header.Size = new System.Drawing.Size(1038, 116);
            this.pnl_rpt_option_header.TabIndex = 8;
            // 
            // pnl_rpt_option
            // 
            this.pnl_rpt_option.BackColor = System.Drawing.Color.DarkKhaki;
            this.pnl_rpt_option.Controls.Add(this.panel2);
            this.pnl_rpt_option.Controls.Add(this.grp_options);
            this.pnl_rpt_option.Controls.Add(this.panel1);
            this.pnl_rpt_option.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_rpt_option.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pnl_rpt_option.Location = new System.Drawing.Point(0, 0);
            this.pnl_rpt_option.Name = "pnl_rpt_option";
            this.pnl_rpt_option.Size = new System.Drawing.Size(974, 114);
            this.pnl_rpt_option.TabIndex = 8;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btn_submit);
            this.panel2.Controls.Add(this.btn_clear);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(835, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(148, 114);
            this.panel2.TabIndex = 4;
            // 
            // btn_clear
            // 
            this.btn_clear.BackColor = System.Drawing.Color.Peru;
            this.btn_clear.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_clear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_clear.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btn_clear.Location = new System.Drawing.Point(6, 59);
            this.btn_clear.Name = "btn_clear";
            this.btn_clear.Size = new System.Drawing.Size(139, 43);
            this.btn_clear.TabIndex = 5;
            this.btn_clear.Text = "Clear Fields";
            this.btn_clear.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_clear.UseVisualStyleBackColor = false;
            this.btn_clear.Click += new System.EventHandler(this.btn_clear_Click);
            // 
            // grp_options
            // 
            this.grp_options.Controls.Add(this.cbo_3);
            this.grp_options.Controls.Add(this.lbl_cbo_3);
            this.grp_options.Controls.Add(this.cbo_2);
            this.grp_options.Controls.Add(this.lbl_cbo_2);
            this.grp_options.Controls.Add(this.chk_3);
            this.grp_options.Controls.Add(this.chk_2);
            this.grp_options.Controls.Add(this.chk_1);
            this.grp_options.Controls.Add(this.cbo_1);
            this.grp_options.Controls.Add(this.lbl_cbo_1);
            this.grp_options.Dock = System.Windows.Forms.DockStyle.Left;
            this.grp_options.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grp_options.ForeColor = System.Drawing.SystemColors.Info;
            this.grp_options.Location = new System.Drawing.Point(189, 0);
            this.grp_options.Name = "grp_options";
            this.grp_options.Size = new System.Drawing.Size(646, 114);
            this.grp_options.TabIndex = 3;
            this.grp_options.TabStop = false;
            this.grp_options.Text = "Other Options";
            // 
            // cbo_3
            // 
            this.cbo_3.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_3.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_3.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_3.FormattingEnabled = true;
            this.cbo_3.Location = new System.Drawing.Point(110, 77);
            this.cbo_3.Name = "cbo_3";
            this.cbo_3.Size = new System.Drawing.Size(342, 23);
            this.cbo_3.TabIndex = 8;
            this.cbo_3.SelectedIndexChanged += new System.EventHandler(this.cbo_3_SelectedIndexChanged);
            // 
            // lbl_cbo_3
            // 
            this.lbl_cbo_3.AutoSize = true;
            this.lbl_cbo_3.Location = new System.Drawing.Point(11, 80);
            this.lbl_cbo_3.Name = "lbl_cbo_3";
            this.lbl_cbo_3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_cbo_3.Size = new System.Drawing.Size(48, 15);
            this.lbl_cbo_3.TabIndex = 7;
            this.lbl_cbo_3.Text = "User ID";
            // 
            // cbo_2
            // 
            this.cbo_2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_2.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_2.FormattingEnabled = true;
            this.cbo_2.Location = new System.Drawing.Point(110, 47);
            this.cbo_2.Name = "cbo_2";
            this.cbo_2.Size = new System.Drawing.Size(342, 23);
            this.cbo_2.TabIndex = 6;
            this.cbo_2.SelectedIndexChanged += new System.EventHandler(this.cbo_2_SelectedIndexChanged);
            // 
            // lbl_cbo_2
            // 
            this.lbl_cbo_2.AutoSize = true;
            this.lbl_cbo_2.Location = new System.Drawing.Point(11, 52);
            this.lbl_cbo_2.Name = "lbl_cbo_2";
            this.lbl_cbo_2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_cbo_2.Size = new System.Drawing.Size(48, 15);
            this.lbl_cbo_2.TabIndex = 5;
            this.lbl_cbo_2.Text = "User ID";
            // 
            // chk_3
            // 
            this.chk_3.AutoSize = true;
            this.chk_3.Location = new System.Drawing.Point(465, 78);
            this.chk_3.Name = "chk_3";
            this.chk_3.Size = new System.Drawing.Size(85, 19);
            this.chk_3.TabIndex = 4;
            this.chk_3.Text = "checkbox3";
            this.chk_3.UseVisualStyleBackColor = true;
            this.chk_3.CheckedChanged += new System.EventHandler(this.chk_3_CheckedChanged);
            // 
            // chk_2
            // 
            this.chk_2.AutoSize = true;
            this.chk_2.Location = new System.Drawing.Point(465, 50);
            this.chk_2.Name = "chk_2";
            this.chk_2.Size = new System.Drawing.Size(85, 19);
            this.chk_2.TabIndex = 3;
            this.chk_2.Text = "checkbox2";
            this.chk_2.UseVisualStyleBackColor = true;
            this.chk_2.CheckedChanged += new System.EventHandler(this.chk_2_CheckedChanged);
            // 
            // chk_1
            // 
            this.chk_1.AutoSize = true;
            this.chk_1.Location = new System.Drawing.Point(465, 22);
            this.chk_1.Name = "chk_1";
            this.chk_1.Size = new System.Drawing.Size(85, 19);
            this.chk_1.TabIndex = 2;
            this.chk_1.Text = "checkbox1";
            this.chk_1.UseVisualStyleBackColor = true;
            this.chk_1.CheckedChanged += new System.EventHandler(this.chk_1_CheckedChanged);
            // 
            // cbo_1
            // 
            this.cbo_1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_1.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_1.FormattingEnabled = true;
            this.cbo_1.Location = new System.Drawing.Point(110, 18);
            this.cbo_1.Name = "cbo_1";
            this.cbo_1.Size = new System.Drawing.Size(342, 23);
            this.cbo_1.TabIndex = 1;
            this.cbo_1.SelectedIndexChanged += new System.EventHandler(this.cbo_1_SelectedIndexChanged);
            // 
            // lbl_cbo_1
            // 
            this.lbl_cbo_1.AutoSize = true;
            this.lbl_cbo_1.Location = new System.Drawing.Point(11, 23);
            this.lbl_cbo_1.Name = "lbl_cbo_1";
            this.lbl_cbo_1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lbl_cbo_1.Size = new System.Drawing.Size(48, 15);
            this.lbl_cbo_1.TabIndex = 1;
            this.lbl_cbo_1.Text = "User ID";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.grp_bybranch);
            this.panel1.Controls.Add(this.grp_bydate);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(189, 114);
            this.panel1.TabIndex = 0;
            // 
            // grp_bybranch
            // 
            this.grp_bybranch.Controls.Add(this.cbo_4);
            this.grp_bybranch.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.grp_bybranch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grp_bybranch.ForeColor = System.Drawing.SystemColors.Info;
            this.grp_bybranch.Location = new System.Drawing.Point(0, 65);
            this.grp_bybranch.Name = "grp_bybranch";
            this.grp_bybranch.Size = new System.Drawing.Size(189, 49);
            this.grp_bybranch.TabIndex = 7;
            this.grp_bybranch.TabStop = false;
            this.grp_bybranch.Text = "By Branch";
            this.grp_bybranch.Visible = false;
            // 
            // cbo_4
            // 
            this.cbo_4.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbo_4.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbo_4.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cbo_4.FormattingEnabled = true;
            this.cbo_4.Location = new System.Drawing.Point(6, 16);
            this.cbo_4.Name = "cbo_4";
            this.cbo_4.Size = new System.Drawing.Size(178, 23);
            this.cbo_4.TabIndex = 9;
            this.cbo_4.SelectedIndexChanged += new System.EventHandler(this.cbo_4_SelectedIndexChanged);
            // 
            // grp_bydate
            // 
            this.grp_bydate.Controls.Add(this.lbl_dt_to);
            this.grp_bydate.Controls.Add(this.dtp_to);
            this.grp_bydate.Controls.Add(this.lbl_dt_frm);
            this.grp_bydate.Controls.Add(this.dtp_frm);
            this.grp_bydate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grp_bydate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grp_bydate.ForeColor = System.Drawing.SystemColors.Info;
            this.grp_bydate.Location = new System.Drawing.Point(0, 0);
            this.grp_bydate.Name = "grp_bydate";
            this.grp_bydate.Size = new System.Drawing.Size(189, 114);
            this.grp_bydate.TabIndex = 4;
            this.grp_bydate.TabStop = false;
            this.grp_bydate.Text = "By Date";
            // 
            // lbl_dt_to
            // 
            this.lbl_dt_to.AutoSize = true;
            this.lbl_dt_to.Location = new System.Drawing.Point(9, 44);
            this.lbl_dt_to.Name = "lbl_dt_to";
            this.lbl_dt_to.Size = new System.Drawing.Size(21, 15);
            this.lbl_dt_to.TabIndex = 3;
            this.lbl_dt_to.Text = "To";
            // 
            // dtp_to
            // 
            this.dtp_to.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtp_to.Location = new System.Drawing.Point(59, 42);
            this.dtp_to.Name = "dtp_to";
            this.dtp_to.Size = new System.Drawing.Size(100, 21);
            this.dtp_to.TabIndex = 2;
            // 
            // lbl_dt_frm
            // 
            this.lbl_dt_frm.AutoSize = true;
            this.lbl_dt_frm.Location = new System.Drawing.Point(7, 19);
            this.lbl_dt_frm.Name = "lbl_dt_frm";
            this.lbl_dt_frm.Size = new System.Drawing.Size(36, 15);
            this.lbl_dt_frm.TabIndex = 1;
            this.lbl_dt_frm.Text = "From";
            // 
            // dtp_frm
            // 
            this.dtp_frm.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtp_frm.Location = new System.Drawing.Point(59, 15);
            this.dtp_frm.Name = "dtp_frm";
            this.dtp_frm.Size = new System.Drawing.Size(100, 21);
            this.dtp_frm.TabIndex = 0;
            // 
            // btn_minimize
            // 
            this.btn_minimize.BackColor = System.Drawing.Color.Peru;
            this.btn_minimize.Dock = System.Windows.Forms.DockStyle.Right;
            this.btn_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_minimize.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_minimize.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btn_minimize.Location = new System.Drawing.Point(974, 0);
            this.btn_minimize.Name = "btn_minimize";
            this.btn_minimize.Size = new System.Drawing.Size(64, 116);
            this.btn_minimize.TabIndex = 7;
            this.btn_minimize.Text = "<<< Hide Options";
            this.btn_minimize.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_minimize.UseVisualStyleBackColor = false;
            this.btn_minimize.Click += new System.EventHandler(this.btn_minimize_Click);
            // 
            // pnl_pbar
            // 
            this.pnl_pbar.Controls.Add(this.textBox1);
            this.pnl_pbar.Controls.Add(this.progressBar1);
            this.pnl_pbar.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnl_pbar.Location = new System.Drawing.Point(0, 116);
            this.pnl_pbar.Name = "pnl_pbar";
            this.pnl_pbar.Size = new System.Drawing.Size(1038, 45);
            this.pnl_pbar.TabIndex = 9;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Info;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(0, 7);
            this.textBox1.Margin = new System.Windows.Forms.Padding(10);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(1038, 15);
            this.textBox1.TabIndex = 8;
            this.textBox1.Text = "Processing. Thank you for your patience.";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 22);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1038, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // crptviewer
            // 
            this.crptviewer.ActiveViewIndex = -1;
            this.crptviewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.crptviewer.Cursor = System.Windows.Forms.Cursors.Default;
            this.crptviewer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.crptviewer.Location = new System.Drawing.Point(0, 161);
            this.crptviewer.Name = "crptviewer";
            this.crptviewer.Padding = new System.Windows.Forms.Padding(10);
            this.crptviewer.SelectionFormula = "";
            this.crptviewer.ShowLogo = false;
            this.crptviewer.Size = new System.Drawing.Size(1038, 422);
            this.crptviewer.TabIndex = 10;
            this.crptviewer.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;
            this.crptviewer.ToolPanelWidth = 210;
            this.crptviewer.ViewTimeSelectionFormula = "";
            // 
            // btn_submit
            // 
            this.btn_submit.BackColor = System.Drawing.Color.SaddleBrown;
            this.btn_submit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_submit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_submit.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.btn_submit.Image = global::Human_Resource_Information_System.Properties.Resources.Print;
            this.btn_submit.Location = new System.Drawing.Point(6, 10);
            this.btn_submit.Name = "btn_submit";
            this.btn_submit.Size = new System.Drawing.Size(139, 43);
            this.btn_submit.TabIndex = 1;
            this.btn_submit.Text = "Preview";
            this.btn_submit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_submit.UseVisualStyleBackColor = false;
            this.btn_submit.Click += new System.EventHandler(this.btn_submit_Click);
            // 
            // RPT_RES_entry
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(1038, 583);
            this.Controls.Add(this.crptviewer);
            this.Controls.Add(this.pnl_pbar);
            this.Controls.Add(this.pnl_rpt_option_header);
            this.Name = "RPT_RES_entry";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "RPT_RES_entry";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.RPT_RES_entry_Load);
            this.pnl_rpt_option_header.ResumeLayout(false);
            this.pnl_rpt_option.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.grp_options.ResumeLayout(false);
            this.grp_options.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.grp_bybranch.ResumeLayout(false);
            this.grp_bydate.ResumeLayout(false);
            this.grp_bydate.PerformLayout();
            this.pnl_pbar.ResumeLayout(false);
            this.pnl_pbar.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.ComponentModel.BackgroundWorker bgworker;
        private System.Windows.Forms.Panel pnl_rpt_option_header;
        private System.Windows.Forms.Panel pnl_rpt_option;
        private System.Windows.Forms.Button btn_minimize;
        private System.Windows.Forms.Panel pnl_pbar;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private CrystalDecisions.Windows.Forms.CrystalReportViewer crptviewer;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btn_submit;
        private System.Windows.Forms.Button btn_clear;
        private System.Windows.Forms.GroupBox grp_options;
        private System.Windows.Forms.ComboBox cbo_3;
        private System.Windows.Forms.Label lbl_cbo_3;
        private System.Windows.Forms.ComboBox cbo_2;
        private System.Windows.Forms.Label lbl_cbo_2;
        private System.Windows.Forms.CheckBox chk_3;
        private System.Windows.Forms.CheckBox chk_2;
        private System.Windows.Forms.CheckBox chk_1;
        private System.Windows.Forms.ComboBox cbo_1;
        private System.Windows.Forms.Label lbl_cbo_1;
        private System.Windows.Forms.GroupBox grp_bybranch;
        private System.Windows.Forms.ComboBox cbo_4;
        private System.Windows.Forms.GroupBox grp_bydate;
        private System.Windows.Forms.Label lbl_dt_to;
        private System.Windows.Forms.DateTimePicker dtp_to;
        private System.Windows.Forms.Label lbl_dt_frm;
        private System.Windows.Forms.DateTimePicker dtp_frm;
    }
}