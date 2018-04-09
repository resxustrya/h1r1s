namespace Human_Resource_Information_System
{
    partial class activation_form
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
            this.txt_key = new System.Windows.Forms.TextBox();
            this.btn_activate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txt_key
            // 
            this.txt_key.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_key.Location = new System.Drawing.Point(72, 38);
            this.txt_key.Name = "txt_key";
            this.txt_key.Size = new System.Drawing.Size(406, 31);
            this.txt_key.TabIndex = 0;
            // 
            // btn_activate
            // 
            this.btn_activate.BackColor = System.Drawing.SystemColors.Info;
            this.btn_activate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_activate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_activate.Location = new System.Drawing.Point(216, 86);
            this.btn_activate.Name = "btn_activate";
            this.btn_activate.Size = new System.Drawing.Size(139, 40);
            this.btn_activate.TabIndex = 1;
            this.btn_activate.Text = "Activate";
            this.btn_activate.UseVisualStyleBackColor = false;
            this.btn_activate.Click += new System.EventHandler(this.btn_activate_Click);
            this.btn_activate.KeyDown += new System.Windows.Forms.KeyEventHandler(this.btn_activate_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.DarkRed;
            this.label1.Location = new System.Drawing.Point(80, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(393, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "*Enter valid license key to activate and use this system in this unit.";
            // 
            // activation_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(536, 148);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_activate);
            this.Controls.Add(this.txt_key);
            this.Name = "activation_form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "activation_form";
            this.Load += new System.EventHandler(this.activation_form_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_key;
        private System.Windows.Forms.Button btn_activate;
        private System.Windows.Forms.Label label1;
    }
}