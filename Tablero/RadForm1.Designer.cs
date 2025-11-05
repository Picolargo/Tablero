namespace Tablero
{
    partial class Editar
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
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition1 = new Telerik.WinControls.UI.TableViewDefinition();
            this.radGridView1 = new Telerik.WinControls.UI.RadGridView();
            this.materialExpansionPanel1 = new MaterialSkin.Controls.MaterialExpansionPanel();
            this.txt_user = new MaterialSkin.Controls.MaterialTextBox2();
            this.txt_password = new MaterialSkin.Controls.MaterialTextBox2();
            this.telerikMetroTheme1 = new Telerik.WinControls.Themes.TelerikMetroTheme();
            ((System.ComponentModel.ISupportInitialize)(this.radGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGridView1.MasterTemplate)).BeginInit();
            this.materialExpansionPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // radGridView1
            // 
            this.radGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.radGridView1.Location = new System.Drawing.Point(0, 0);
            // 
            // 
            // 
            this.radGridView1.MasterTemplate.AllowAddNewRow = false;
            this.radGridView1.MasterTemplate.AllowDeleteRow = false;
            this.radGridView1.MasterTemplate.AllowEditRow = false;
            this.radGridView1.MasterTemplate.AutoSizeColumnsMode = Telerik.WinControls.UI.GridViewAutoSizeColumnsMode.Fill;
            this.radGridView1.MasterTemplate.ViewDefinition = tableViewDefinition1;
            this.radGridView1.Name = "radGridView1";
            this.radGridView1.ReadOnly = true;
            this.radGridView1.Size = new System.Drawing.Size(977, 618);
            this.radGridView1.TabIndex = 4;
            this.radGridView1.ThemeName = "TelerikMetro";
            this.radGridView1.Visible = false;
            this.radGridView1.CellDoubleClick += new Telerik.WinControls.UI.GridViewCellEventHandler(this.radGridView1_CellDoubleClick);
            // 
            // materialExpansionPanel1
            // 
            this.materialExpansionPanel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialExpansionPanel1.CancelButtonText = "CANCELAR";
            this.materialExpansionPanel1.Controls.Add(this.txt_user);
            this.materialExpansionPanel1.Controls.Add(this.txt_password);
            this.materialExpansionPanel1.Depth = 0;
            this.materialExpansionPanel1.Description = "Acceso solo para administrador.";
            this.materialExpansionPanel1.ExpandHeight = 283;
            this.materialExpansionPanel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialExpansionPanel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialExpansionPanel1.Location = new System.Drawing.Point(236, 124);
            this.materialExpansionPanel1.Margin = new System.Windows.Forms.Padding(16);
            this.materialExpansionPanel1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialExpansionPanel1.Name = "materialExpansionPanel1";
            this.materialExpansionPanel1.Padding = new System.Windows.Forms.Padding(24, 66, 26, 18);
            this.materialExpansionPanel1.ShowCollapseExpand = false;
            this.materialExpansionPanel1.Size = new System.Drawing.Size(518, 283);
            this.materialExpansionPanel1.TabIndex = 3;
            this.materialExpansionPanel1.Title = "Validación de Seguridad";
            this.materialExpansionPanel1.ValidationButtonEnable = true;
            this.materialExpansionPanel1.ValidationButtonText = "INICIAR";
            // 
            // txt_user
            // 
            this.txt_user.AnimateReadOnly = false;
            this.txt_user.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.txt_user.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
            this.txt_user.Depth = 0;
            this.txt_user.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_user.HelperText = "Escriba el usuario administrador.";
            this.txt_user.HideSelection = true;
            this.txt_user.Hint = "Usuario";
            this.txt_user.LeadingIcon = global::Tablero.Properties.Resources._1564535_customer_user_userphoto_account_person_icon;
            this.txt_user.Location = new System.Drawing.Point(27, 69);
            this.txt_user.MaxLength = 32767;
            this.txt_user.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_user.Name = "txt_user";
            this.txt_user.PasswordChar = '\0';
            this.txt_user.PrefixSuffixText = null;
            this.txt_user.ReadOnly = false;
            this.txt_user.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txt_user.SelectedText = "";
            this.txt_user.SelectionLength = 0;
            this.txt_user.SelectionStart = 0;
            this.txt_user.ShortcutsEnabled = true;
            this.txt_user.ShowAssistiveText = true;
            this.txt_user.Size = new System.Drawing.Size(250, 64);
            this.txt_user.TabIndex = 1;
            this.txt_user.TabStop = false;
            this.txt_user.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.txt_user.TrailingIcon = null;
            this.txt_user.UseSystemPasswordChar = false;
            // 
            // txt_password
            // 
            this.txt_password.AnimateReadOnly = false;
            this.txt_password.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.txt_password.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
            this.txt_password.Depth = 0;
            this.txt_password.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_password.HelperText = "Su clave es confidencial.";
            this.txt_password.HideSelection = true;
            this.txt_password.Hint = "Contraseña";
            this.txt_password.LeadingIcon = global::Tablero.Properties.Resources._49855_closed_padlock_icon;
            this.txt_password.Location = new System.Drawing.Point(27, 139);
            this.txt_password.MaxLength = 32767;
            this.txt_password.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_password.Name = "txt_password";
            this.txt_password.PasswordChar = '*';
            this.txt_password.PrefixSuffixText = null;
            this.txt_password.ReadOnly = false;
            this.txt_password.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txt_password.SelectedText = "";
            this.txt_password.SelectionLength = 0;
            this.txt_password.SelectionStart = 0;
            this.txt_password.ShortcutsEnabled = true;
            this.txt_password.ShowAssistiveText = true;
            this.txt_password.Size = new System.Drawing.Size(250, 64);
            this.txt_password.TabIndex = 2;
            this.txt_password.TabStop = false;
            this.txt_password.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.txt_password.TrailingIcon = null;
            this.txt_password.UseSystemPasswordChar = false;
            // 
            // Editar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(977, 618);
            this.Controls.Add(this.materialExpansionPanel1);
            this.Controls.Add(this.radGridView1);
            this.Name = "Editar";
            this.Text = "Fichas";
            this.ThemeName = "TelerikMetro";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Editar_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.radGridView1.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGridView1)).EndInit();
            this.materialExpansionPanel1.ResumeLayout(false);
            this.materialExpansionPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Telerik.WinControls.UI.RadGridView radGridView1;
        private Telerik.WinControls.Themes.TelerikMetroTheme telerikMetroTheme1;
        private MaterialSkin.Controls.MaterialExpansionPanel materialExpansionPanel1;
        private MaterialSkin.Controls.MaterialTextBox2 txt_user;
        private MaterialSkin.Controls.MaterialTextBox2 txt_password;
    }
}
