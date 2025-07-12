namespace Tablero
{
    partial class Form_principal
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_principal));
            this.materialTabControl1 = new MaterialSkin.Controls.MaterialTabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.materialTabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // materialTabControl1
            // 
            this.materialTabControl1.Controls.Add(this.tabPage1);
            this.materialTabControl1.Controls.Add(this.tabPage2);
            this.materialTabControl1.Depth = 0;
            this.materialTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.materialTabControl1.ImageList = this.imageList1;
            this.materialTabControl1.Location = new System.Drawing.Point(3, 64);
            this.materialTabControl1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialTabControl1.Multiline = true;
            this.materialTabControl1.Name = "materialTabControl1";
            this.materialTabControl1.SelectedIndex = 0;
            this.materialTabControl1.Size = new System.Drawing.Size(1263, 629);
            this.materialTabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1255, 603);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(192, 74);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "excel_microsoft_icon");
            this.imageList1.Images.SetKeyName(1, "clipboard_list_checklist_document_check_icon");
            this.imageList1.Images.SetKeyName(2, "id_card_person_identity_business_icon");
            this.imageList1.Images.SetKeyName(3, "white_board_education_business_presentation_icon");
            this.imageList1.Images.SetKeyName(4, "computer_technology_business_digital_monitor_icon");
            this.imageList1.Images.SetKeyName(5, "floppy_disc_save_storage_disk_icon");
            this.imageList1.Images.SetKeyName(6, "arrow_entrance_exit_internet_log_icon");
            this.imageList1.Images.SetKeyName(7, "arrow_entrance_in_internet_log_icon");
            this.imageList1.Images.SetKeyName(8, "check_document_file_internet_report_icon");
            this.imageList1.Images.SetKeyName(9, "app_interface_internet_profile_ui_icon");
            this.imageList1.Images.SetKeyName(10, "internet_locked_padlock_password_secure_icon");
            this.imageList1.Images.SetKeyName(11, "internet_lock_locked_padlock_password_icon");
            this.imageList1.Images.SetKeyName(12, "internet_lock_locked_padlock_password_icon");
            this.imageList1.Images.SetKeyName(13, "door_internet_key_lock_password_icon");
            this.imageList1.Images.SetKeyName(14, "disable_eye_hidden_hide_internet_icon");
            this.imageList1.Images.SetKeyName(15, "business_eye_focus_internet_security_icon");
            this.imageList1.Images.SetKeyName(16, "eye_focus_internet_scan_security_icon");
            this.imageList1.Images.SetKeyName(17, "home_house_internet_locked_padlock_icon");
            this.imageList1.Images.SetKeyName(18, "home_house_internet_lock_locked_icon");
            this.imageList1.Images.SetKeyName(19, "engine_gear_internet_option_security_icon");
            this.imageList1.Images.SetKeyName(20, "home_house_internet_network_security_icon");
            // 
            // Form_principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1269, 696);
            this.Controls.Add(this.materialTabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form_principal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.materialTabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private MaterialSkin.Controls.MaterialTabControl materialTabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ImageList imageList1;
    }
}

