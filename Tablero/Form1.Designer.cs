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
            this.materialCard3 = new MaterialSkin.Controls.MaterialCard();
            this.Card_verde = new MaterialSkin.Controls.MaterialCard();
            this.pnl_principal = new System.Windows.Forms.Panel();
            this.txt_meta_p = new MaterialSkin.Controls.MaterialTextBox();
            this.Mask_txt_hr1 = new MaterialSkin.Controls.MaterialMaskedTextBox();
            this.Mask_txt_hr2 = new MaterialSkin.Controls.MaterialMaskedTextBox();
            this.txt_hr_disponible = new MaterialSkin.Controls.MaterialTextBox();
            this.materialCard4 = new MaterialSkin.Controls.MaterialCard();
            this.lbl_meta2 = new MaterialSkin.Controls.MaterialLabel();
            this.lbl_meta = new MaterialSkin.Controls.MaterialLabel();
            this.cb_OP = new MaterialSkin.Controls.MaterialComboBox();
            this.materialCard2 = new MaterialSkin.Controls.MaterialCard();
            this.cb_Turno = new MaterialSkin.Controls.MaterialComboBox();
            this.cb_Area = new MaterialSkin.Controls.MaterialComboBox();
            this.dtp1 = new System.Windows.Forms.DateTimePicker();
            this.materialCard1 = new MaterialSkin.Controls.MaterialCard();
            this.lbl_nom2 = new MaterialSkin.Controls.MaterialLabel();
            this.lbl_no_emp2 = new MaterialSkin.Controls.MaterialLabel();
            this.lbl_Nom = new MaterialSkin.Controls.MaterialLabel();
            this.lbl_user_no_emp = new MaterialSkin.Controls.MaterialLabel();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.materialTextBox1 = new MaterialSkin.Controls.MaterialTextBox();
            this.pnl_Metahora = new System.Windows.Forms.Panel();
            this.materialTabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.materialCard3.SuspendLayout();
            this.pnl_principal.SuspendLayout();
            this.materialCard4.SuspendLayout();
            this.materialCard2.SuspendLayout();
            this.materialCard1.SuspendLayout();
            this.pnl_Metahora.SuspendLayout();
            this.SuspendLayout();
            // 
            // materialTabControl1
            // 
            this.materialTabControl1.Controls.Add(this.tabPage1);
            this.materialTabControl1.Controls.Add(this.tabPage2);
            this.materialTabControl1.Controls.Add(this.tabPage3);
            this.materialTabControl1.Controls.Add(this.tabPage4);
            this.materialTabControl1.Depth = 0;
            this.materialTabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.materialTabControl1.ImageList = this.imageList1;
            this.materialTabControl1.Location = new System.Drawing.Point(3, 64);
            this.materialTabControl1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialTabControl1.Multiline = true;
            this.materialTabControl1.Name = "materialTabControl1";
            this.materialTabControl1.SelectedIndex = 0;
            this.materialTabControl1.Size = new System.Drawing.Size(1263, 939);
            this.materialTabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.materialCard3);
            this.tabPage1.Controls.Add(this.materialCard2);
            this.tabPage1.Controls.Add(this.materialCard1);
            this.tabPage1.ImageKey = "home_house";
            this.tabPage1.Location = new System.Drawing.Point(4, 39);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1255, 896);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Home";
            // 
            // materialCard3
            // 
            this.materialCard3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard3.Controls.Add(this.pnl_principal);
            this.materialCard3.Controls.Add(this.Card_verde);
            this.materialCard3.Controls.Add(this.materialCard4);
            this.materialCard3.Controls.Add(this.lbl_meta);
            this.materialCard3.Controls.Add(this.cb_OP);
            this.materialCard3.Depth = 0;
            this.materialCard3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard3.Location = new System.Drawing.Point(14, 329);
            this.materialCard3.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard3.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard3.Name = "materialCard3";
            this.materialCard3.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard3.Size = new System.Drawing.Size(651, 458);
            this.materialCard3.TabIndex = 3;
            // 
            // Card_verde
            // 
            this.Card_verde.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.Card_verde.Depth = 0;
            this.Card_verde.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Card_verde.Location = new System.Drawing.Point(18, 80);
            this.Card_verde.Margin = new System.Windows.Forms.Padding(14);
            this.Card_verde.MouseState = MaterialSkin.MouseState.HOVER;
            this.Card_verde.Name = "Card_verde";
            this.Card_verde.Padding = new System.Windows.Forms.Padding(14);
            this.Card_verde.Size = new System.Drawing.Size(369, 215);
            this.Card_verde.TabIndex = 4;
            // 
            // pnl_principal
            // 
            this.pnl_principal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.pnl_principal.Controls.Add(this.materialTextBox1);
            this.pnl_principal.Controls.Add(this.txt_meta_p);
            this.pnl_principal.Controls.Add(this.Mask_txt_hr1);
            this.pnl_principal.Controls.Add(this.Mask_txt_hr2);
            this.pnl_principal.Controls.Add(this.txt_hr_disponible);
            this.pnl_principal.Location = new System.Drawing.Point(18, 69);
            this.pnl_principal.Name = "pnl_principal";
            this.pnl_principal.Size = new System.Drawing.Size(369, 226);
            this.pnl_principal.TabIndex = 5;
            // 
            // txt_meta_p
            // 
            this.txt_meta_p.AnimateReadOnly = false;
            this.txt_meta_p.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_meta_p.Depth = 0;
            this.txt_meta_p.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_meta_p.Hint = "Meta programada";
            this.txt_meta_p.LeadingIcon = null;
            this.txt_meta_p.Location = new System.Drawing.Point(195, 86);
            this.txt_meta_p.MaxLength = 50;
            this.txt_meta_p.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_meta_p.Multiline = false;
            this.txt_meta_p.Name = "txt_meta_p";
            this.txt_meta_p.Size = new System.Drawing.Size(157, 50);
            this.txt_meta_p.TabIndex = 3;
            this.txt_meta_p.Text = "";
            this.txt_meta_p.TrailingIcon = null;
            // 
            // Mask_txt_hr1
            // 
            this.Mask_txt_hr1.AllowPromptAsInput = true;
            this.Mask_txt_hr1.AnimateReadOnly = false;
            this.Mask_txt_hr1.AsciiOnly = false;
            this.Mask_txt_hr1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.Mask_txt_hr1.BeepOnError = false;
            this.Mask_txt_hr1.CutCopyMaskFormat = System.Windows.Forms.MaskFormat.IncludeLiterals;
            this.Mask_txt_hr1.Depth = 0;
            this.Mask_txt_hr1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.Mask_txt_hr1.HidePromptOnLeave = false;
            this.Mask_txt_hr1.HideSelection = true;
            this.Mask_txt_hr1.Hint = "Hora inicio real";
            this.Mask_txt_hr1.InsertKeyMode = System.Windows.Forms.InsertKeyMode.Default;
            this.Mask_txt_hr1.LeadingIcon = null;
            this.Mask_txt_hr1.Location = new System.Drawing.Point(17, 11);
            this.Mask_txt_hr1.Mask = "00:00";
            this.Mask_txt_hr1.MaxLength = 32767;
            this.Mask_txt_hr1.MouseState = MaterialSkin.MouseState.OUT;
            this.Mask_txt_hr1.Name = "Mask_txt_hr1";
            this.Mask_txt_hr1.PasswordChar = '\0';
            this.Mask_txt_hr1.PrefixSuffixText = null;
            this.Mask_txt_hr1.PromptChar = '_';
            this.Mask_txt_hr1.ReadOnly = false;
            this.Mask_txt_hr1.RejectInputOnFirstFailure = false;
            this.Mask_txt_hr1.ResetOnPrompt = true;
            this.Mask_txt_hr1.ResetOnSpace = true;
            this.Mask_txt_hr1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Mask_txt_hr1.SelectedText = "";
            this.Mask_txt_hr1.SelectionLength = 0;
            this.Mask_txt_hr1.SelectionStart = 0;
            this.Mask_txt_hr1.ShortcutsEnabled = true;
            this.Mask_txt_hr1.Size = new System.Drawing.Size(157, 48);
            this.Mask_txt_hr1.SkipLiterals = true;
            this.Mask_txt_hr1.TabIndex = 0;
            this.Mask_txt_hr1.TabStop = false;
            this.Mask_txt_hr1.Text = "  :";
            this.Mask_txt_hr1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Mask_txt_hr1.TextMaskFormat = System.Windows.Forms.MaskFormat.IncludeLiterals;
            this.Mask_txt_hr1.TrailingIcon = null;
            this.Mask_txt_hr1.UseSystemPasswordChar = false;
            this.Mask_txt_hr1.ValidatingType = null;
            // 
            // Mask_txt_hr2
            // 
            this.Mask_txt_hr2.AllowPromptAsInput = true;
            this.Mask_txt_hr2.AnimateReadOnly = false;
            this.Mask_txt_hr2.AsciiOnly = false;
            this.Mask_txt_hr2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.Mask_txt_hr2.BeepOnError = false;
            this.Mask_txt_hr2.CutCopyMaskFormat = System.Windows.Forms.MaskFormat.IncludeLiterals;
            this.Mask_txt_hr2.Depth = 0;
            this.Mask_txt_hr2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.Mask_txt_hr2.HidePromptOnLeave = false;
            this.Mask_txt_hr2.HideSelection = true;
            this.Mask_txt_hr2.Hint = "Hora final real";
            this.Mask_txt_hr2.InsertKeyMode = System.Windows.Forms.InsertKeyMode.Default;
            this.Mask_txt_hr2.LeadingIcon = null;
            this.Mask_txt_hr2.Location = new System.Drawing.Point(195, 11);
            this.Mask_txt_hr2.Mask = "00:00";
            this.Mask_txt_hr2.MaxLength = 32767;
            this.Mask_txt_hr2.MouseState = MaterialSkin.MouseState.OUT;
            this.Mask_txt_hr2.Name = "Mask_txt_hr2";
            this.Mask_txt_hr2.PasswordChar = '\0';
            this.Mask_txt_hr2.PrefixSuffixText = null;
            this.Mask_txt_hr2.PromptChar = '_';
            this.Mask_txt_hr2.ReadOnly = false;
            this.Mask_txt_hr2.RejectInputOnFirstFailure = false;
            this.Mask_txt_hr2.ResetOnPrompt = true;
            this.Mask_txt_hr2.ResetOnSpace = true;
            this.Mask_txt_hr2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Mask_txt_hr2.SelectedText = "";
            this.Mask_txt_hr2.SelectionLength = 0;
            this.Mask_txt_hr2.SelectionStart = 0;
            this.Mask_txt_hr2.ShortcutsEnabled = true;
            this.Mask_txt_hr2.Size = new System.Drawing.Size(157, 48);
            this.Mask_txt_hr2.SkipLiterals = true;
            this.Mask_txt_hr2.TabIndex = 1;
            this.Mask_txt_hr2.TabStop = false;
            this.Mask_txt_hr2.Text = "  :";
            this.Mask_txt_hr2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Mask_txt_hr2.TextMaskFormat = System.Windows.Forms.MaskFormat.IncludeLiterals;
            this.Mask_txt_hr2.TrailingIcon = null;
            this.Mask_txt_hr2.UseSystemPasswordChar = false;
            this.Mask_txt_hr2.ValidatingType = null;
            // 
            // txt_hr_disponible
            // 
            this.txt_hr_disponible.AnimateReadOnly = false;
            this.txt_hr_disponible.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_hr_disponible.Depth = 0;
            this.txt_hr_disponible.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_hr_disponible.Hint = "Horas disponibles";
            this.txt_hr_disponible.LeadingIcon = null;
            this.txt_hr_disponible.Location = new System.Drawing.Point(17, 86);
            this.txt_hr_disponible.MaxLength = 50;
            this.txt_hr_disponible.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_hr_disponible.Multiline = false;
            this.txt_hr_disponible.Name = "txt_hr_disponible";
            this.txt_hr_disponible.Size = new System.Drawing.Size(157, 50);
            this.txt_hr_disponible.TabIndex = 2;
            this.txt_hr_disponible.Text = "";
            this.txt_hr_disponible.TrailingIcon = null;
            // 
            // materialCard4
            // 
            this.materialCard4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard4.Controls.Add(this.pnl_Metahora);
            this.materialCard4.Depth = 0;
            this.materialCard4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard4.Location = new System.Drawing.Point(543, 13);
            this.materialCard4.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard4.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard4.Name = "materialCard4";
            this.materialCard4.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard4.Size = new System.Drawing.Size(98, 46);
            this.materialCard4.TabIndex = 3;
            // 
            // lbl_meta2
            // 
            this.lbl_meta2.AutoSize = true;
            this.lbl_meta2.Depth = 0;
            this.lbl_meta2.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_meta2.Location = new System.Drawing.Point(17, 13);
            this.lbl_meta2.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_meta2.Name = "lbl_meta2";
            this.lbl_meta2.Size = new System.Drawing.Size(41, 19);
            this.lbl_meta2.TabIndex = 0;
            this.lbl_meta2.Text = "          ";
            // 
            // lbl_meta
            // 
            this.lbl_meta.AutoSize = true;
            this.lbl_meta.Depth = 0;
            this.lbl_meta.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_meta.Location = new System.Drawing.Point(430, 27);
            this.lbl_meta.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_meta.Name = "lbl_meta";
            this.lbl_meta.Size = new System.Drawing.Size(100, 19);
            this.lbl_meta.TabIndex = 2;
            this.lbl_meta.Text = "Meta por hora";
            // 
            // cb_OP
            // 
            this.cb_OP.AutoResize = false;
            this.cb_OP.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.cb_OP.Depth = 0;
            this.cb_OP.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.cb_OP.DropDownHeight = 174;
            this.cb_OP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_OP.DropDownWidth = 121;
            this.cb_OP.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            this.cb_OP.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.cb_OP.FormattingEnabled = true;
            this.cb_OP.Hint = "OP";
            this.cb_OP.IntegralHeight = false;
            this.cb_OP.ItemHeight = 43;
            this.cb_OP.Location = new System.Drawing.Point(18, 14);
            this.cb_OP.MaxDropDownItems = 4;
            this.cb_OP.MouseState = MaterialSkin.MouseState.OUT;
            this.cb_OP.Name = "cb_OP";
            this.cb_OP.Size = new System.Drawing.Size(259, 49);
            this.cb_OP.StartIndex = 0;
            this.cb_OP.TabIndex = 1;
            // 
            // materialCard2
            // 
            this.materialCard2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard2.Controls.Add(this.cb_Turno);
            this.materialCard2.Controls.Add(this.cb_Area);
            this.materialCard2.Controls.Add(this.dtp1);
            this.materialCard2.Depth = 0;
            this.materialCard2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard2.Location = new System.Drawing.Point(14, 123);
            this.materialCard2.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard2.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard2.Name = "materialCard2";
            this.materialCard2.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard2.Size = new System.Drawing.Size(498, 178);
            this.materialCard2.TabIndex = 2;
            // 
            // cb_Turno
            // 
            this.cb_Turno.AutoResize = false;
            this.cb_Turno.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.cb_Turno.Depth = 0;
            this.cb_Turno.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.cb_Turno.DropDownHeight = 174;
            this.cb_Turno.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_Turno.DropDownWidth = 121;
            this.cb_Turno.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            this.cb_Turno.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.cb_Turno.FormattingEnabled = true;
            this.cb_Turno.Hint = "Turno";
            this.cb_Turno.IntegralHeight = false;
            this.cb_Turno.ItemHeight = 43;
            this.cb_Turno.Items.AddRange(new object[] {
            "1",
            "2",
            "3"});
            this.cb_Turno.Location = new System.Drawing.Point(49, 112);
            this.cb_Turno.MaxDropDownItems = 4;
            this.cb_Turno.MouseState = MaterialSkin.MouseState.OUT;
            this.cb_Turno.Name = "cb_Turno";
            this.cb_Turno.Size = new System.Drawing.Size(259, 49);
            this.cb_Turno.StartIndex = -1;
            this.cb_Turno.TabIndex = 5;
            // 
            // cb_Area
            // 
            this.cb_Area.AutoResize = false;
            this.cb_Area.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.cb_Area.Depth = 0;
            this.cb_Area.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.cb_Area.DropDownHeight = 174;
            this.cb_Area.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_Area.DropDownWidth = 121;
            this.cb_Area.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            this.cb_Area.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.cb_Area.FormattingEnabled = true;
            this.cb_Area.Hint = "Área";
            this.cb_Area.IntegralHeight = false;
            this.cb_Area.ItemHeight = 43;
            this.cb_Area.Items.AddRange(new object[] {
            "TUNEL/ SUMERGIDOR",
            "Despegue",
            "Polvos",
            "Evaporado",
            "Grind",
            "Jalapeños",
            "Inspeccion 1",
            "Inspeccion 2",
            "Inspeccion y Empacado"});
            this.cb_Area.Location = new System.Drawing.Point(49, 11);
            this.cb_Area.MaxDropDownItems = 4;
            this.cb_Area.MouseState = MaterialSkin.MouseState.OUT;
            this.cb_Area.Name = "cb_Area";
            this.cb_Area.Size = new System.Drawing.Size(259, 49);
            this.cb_Area.StartIndex = -1;
            this.cb_Area.TabIndex = 3;
            // 
            // dtp1
            // 
            this.dtp1.Location = new System.Drawing.Point(49, 76);
            this.dtp1.Name = "dtp1";
            this.dtp1.Size = new System.Drawing.Size(259, 20);
            this.dtp1.TabIndex = 1;
            // 
            // materialCard1
            // 
            this.materialCard1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard1.Controls.Add(this.lbl_nom2);
            this.materialCard1.Controls.Add(this.lbl_no_emp2);
            this.materialCard1.Controls.Add(this.lbl_Nom);
            this.materialCard1.Controls.Add(this.lbl_user_no_emp);
            this.materialCard1.Depth = 0;
            this.materialCard1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard1.Location = new System.Drawing.Point(17, 17);
            this.materialCard1.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard1.Name = "materialCard1";
            this.materialCard1.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard1.Size = new System.Drawing.Size(498, 78);
            this.materialCard1.TabIndex = 1;
            // 
            // lbl_nom2
            // 
            this.lbl_nom2.AutoSize = true;
            this.lbl_nom2.BackColor = System.Drawing.Color.Transparent;
            this.lbl_nom2.Depth = 0;
            this.lbl_nom2.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_nom2.Location = new System.Drawing.Point(179, 47);
            this.lbl_nom2.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_nom2.Name = "lbl_nom2";
            this.lbl_nom2.Size = new System.Drawing.Size(25, 19);
            this.lbl_nom2.TabIndex = 7;
            this.lbl_nom2.Text = "      ";
            // 
            // lbl_no_emp2
            // 
            this.lbl_no_emp2.AutoSize = true;
            this.lbl_no_emp2.BackColor = System.Drawing.Color.Transparent;
            this.lbl_no_emp2.Depth = 0;
            this.lbl_no_emp2.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_no_emp2.Location = new System.Drawing.Point(114, 14);
            this.lbl_no_emp2.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_no_emp2.Name = "lbl_no_emp2";
            this.lbl_no_emp2.Size = new System.Drawing.Size(25, 19);
            this.lbl_no_emp2.TabIndex = 6;
            this.lbl_no_emp2.Text = "      ";
            // 
            // lbl_Nom
            // 
            this.lbl_Nom.AutoSize = true;
            this.lbl_Nom.Depth = 0;
            this.lbl_Nom.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_Nom.Location = new System.Drawing.Point(12, 47);
            this.lbl_Nom.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_Nom.Name = "lbl_Nom";
            this.lbl_Nom.Size = new System.Drawing.Size(161, 19);
            this.lbl_Nom.TabIndex = 3;
            this.lbl_Nom.Text = "Nombre del Supervisor";
            // 
            // lbl_user_no_emp
            // 
            this.lbl_user_no_emp.AutoSize = true;
            this.lbl_user_no_emp.Depth = 0;
            this.lbl_user_no_emp.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_user_no_emp.Location = new System.Drawing.Point(12, 14);
            this.lbl_user_no_emp.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_user_no_emp.Name = "lbl_user_no_emp";
            this.lbl_user_no_emp.Size = new System.Drawing.Size(96, 19);
            this.lbl_user_no_emp.TabIndex = 2;
            this.lbl_user_no_emp.Text = "No Empleado";
            // 
            // tabPage2
            // 
            this.tabPage2.ImageKey = "chart_graph_statistics_bar";
            this.tabPage2.Location = new System.Drawing.Point(4, 39);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1255, 896);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Graficas";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.ImageKey = "group_users_people";
            this.tabPage3.Location = new System.Drawing.Point(4, 39);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(1255, 896);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Usuarios";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.ImageKey = "gear_fill";
            this.tabPage4.Location = new System.Drawing.Point(4, 39);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(1255, 896);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Configuración";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "gear_fill");
            this.imageList1.Images.SetKeyName(1, "group_users_people");
            this.imageList1.Images.SetKeyName(2, "chart_graph_statistics_bar");
            this.imageList1.Images.SetKeyName(3, "home_house");
            // 
            // materialTextBox1
            // 
            this.materialTextBox1.AnimateReadOnly = false;
            this.materialTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.materialTextBox1.Depth = 0;
            this.materialTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialTextBox1.Hint = "Horas disponibles menos tiempo muerto";
            this.materialTextBox1.LeadingIcon = null;
            this.materialTextBox1.Location = new System.Drawing.Point(17, 156);
            this.materialTextBox1.MaxLength = 50;
            this.materialTextBox1.MouseState = MaterialSkin.MouseState.OUT;
            this.materialTextBox1.Multiline = false;
            this.materialTextBox1.Name = "materialTextBox1";
            this.materialTextBox1.Size = new System.Drawing.Size(335, 50);
            this.materialTextBox1.TabIndex = 4;
            this.materialTextBox1.Text = "";
            this.materialTextBox1.TrailingIcon = null;
            // 
            // pnl_Metahora
            // 
            this.pnl_Metahora.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.pnl_Metahora.Controls.Add(this.lbl_meta2);
            this.pnl_Metahora.Location = new System.Drawing.Point(0, 0);
            this.pnl_Metahora.Name = "pnl_Metahora";
            this.pnl_Metahora.Size = new System.Drawing.Size(98, 46);
            this.pnl_Metahora.TabIndex = 0;
            // 
            // Form_principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1269, 1006);
            this.Controls.Add(this.materialTabControl1);
            this.DrawerShowIconsWhenHidden = true;
            this.DrawerTabControl = this.materialTabControl1;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form_principal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tablero";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.materialTabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.materialCard3.ResumeLayout(false);
            this.materialCard3.PerformLayout();
            this.pnl_principal.ResumeLayout(false);
            this.materialCard4.ResumeLayout(false);
            this.materialCard2.ResumeLayout(false);
            this.materialCard1.ResumeLayout(false);
            this.materialCard1.PerformLayout();
            this.pnl_Metahora.ResumeLayout(false);
            this.pnl_Metahora.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private MaterialSkin.Controls.MaterialTabControl materialTabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private MaterialSkin.Controls.MaterialCard materialCard1;
        private System.Windows.Forms.DateTimePicker dtp1;
        private MaterialSkin.Controls.MaterialLabel lbl_no_emp2;
        private MaterialSkin.Controls.MaterialLabel lbl_Nom;
        private MaterialSkin.Controls.MaterialLabel lbl_user_no_emp;
        private MaterialSkin.Controls.MaterialCard materialCard2;
        private MaterialSkin.Controls.MaterialLabel lbl_nom2;
        private MaterialSkin.Controls.MaterialComboBox cb_Area;
        private MaterialSkin.Controls.MaterialComboBox cb_Turno;
        private MaterialSkin.Controls.MaterialCard materialCard3;
        private MaterialSkin.Controls.MaterialLabel lbl_meta;
        private MaterialSkin.Controls.MaterialComboBox cb_OP;
        private MaterialSkin.Controls.MaterialCard materialCard4;
        private MaterialSkin.Controls.MaterialLabel lbl_meta2;
        private MaterialSkin.Controls.MaterialMaskedTextBox Mask_txt_hr1;
        private MaterialSkin.Controls.MaterialTextBox txt_hr_disponible;
        private MaterialSkin.Controls.MaterialMaskedTextBox Mask_txt_hr2;
        private MaterialSkin.Controls.MaterialTextBox txt_meta_p;
        private MaterialSkin.Controls.MaterialCard Card_verde;
        private System.Windows.Forms.Panel pnl_principal;
        private MaterialSkin.Controls.MaterialTextBox materialTextBox1;
        private System.Windows.Forms.Panel pnl_Metahora;
    }
}

