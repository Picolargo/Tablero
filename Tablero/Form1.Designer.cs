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
            this.pnl_principal = new System.Windows.Forms.Panel();
            this.txt_hr_Disp_TM = new MaterialSkin.Controls.MaterialTextBox();
            this.txt_meta_p = new MaterialSkin.Controls.MaterialTextBox();
            this.Mask_txt_hr1 = new MaterialSkin.Controls.MaterialMaskedTextBox();
            this.Mask_txt_hr2 = new MaterialSkin.Controls.MaterialMaskedTextBox();
            this.txt_hr_disponible = new MaterialSkin.Controls.MaterialTextBox();
            this.Card_hrs = new MaterialSkin.Controls.MaterialCard();
            this.materialCard4 = new MaterialSkin.Controls.MaterialCard();
            this.pnl_Metahora = new System.Windows.Forms.Panel();
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
            this.materialCard5 = new MaterialSkin.Controls.MaterialCard();
            this.card_datos = new MaterialSkin.Controls.MaterialCard();
            this.txt_perdida = new MaterialSkin.Controls.MaterialTextBox();
            this.txt_ftt = new MaterialSkin.Controls.MaterialTextBox();
            this.txt_kg_ok = new MaterialSkin.Controls.MaterialTextBox();
            this.txt_kg_F_Espec = new MaterialSkin.Controls.MaterialTextBox();
            this.txt_Up_time = new MaterialSkin.Controls.MaterialTextBox();
            this.txt_Produccion = new MaterialSkin.Controls.MaterialTextBox();
            this.Txt_Merma = new MaterialSkin.Controls.MaterialTextBox();
            this.Txt_Personal_Op = new MaterialSkin.Controls.MaterialTextBox();
            this.pnl_ftt = new System.Windows.Forms.Panel();
            this.materialCard6 = new MaterialSkin.Controls.MaterialCard();
            this.txt_Tiempo_comida = new MaterialSkin.Controls.MaterialTextBox();
            this.dgv_mecanico = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.card_dgv_mecanico = new MaterialSkin.Controls.MaterialCard();
            this.txt_TM_mecanico = new MaterialSkin.Controls.MaterialTextBox();
            this.lbl_min_comida = new MaterialSkin.Controls.MaterialLabel();
            this.card_dgv_almacen = new MaterialSkin.Controls.MaterialCard();
            this.dgv_almacen = new System.Windows.Forms.DataGridView();
            this.txt_TM_almacen = new MaterialSkin.Controls.MaterialTextBox();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_TM_operativo = new MaterialSkin.Controls.MaterialTextBox();
            this.card_dgv_operativo = new MaterialSkin.Controls.MaterialCard();
            this.dgv_operativo = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_TM_calidad = new MaterialSkin.Controls.MaterialTextBox();
            this.card_dgv_calidad = new MaterialSkin.Controls.MaterialCard();
            this.dgv_calidad = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.materialTabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.materialCard3.SuspendLayout();
            this.pnl_principal.SuspendLayout();
            this.Card_hrs.SuspendLayout();
            this.materialCard4.SuspendLayout();
            this.pnl_Metahora.SuspendLayout();
            this.materialCard2.SuspendLayout();
            this.materialCard1.SuspendLayout();
            this.materialCard5.SuspendLayout();
            this.card_datos.SuspendLayout();
            this.pnl_ftt.SuspendLayout();
            this.materialCard6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_mecanico)).BeginInit();
            this.card_dgv_mecanico.SuspendLayout();
            this.card_dgv_almacen.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_almacen)).BeginInit();
            this.card_dgv_operativo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_operativo)).BeginInit();
            this.card_dgv_calidad.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_calidad)).BeginInit();
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
            this.materialTabControl1.Size = new System.Drawing.Size(1713, 1026);
            this.materialTabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.materialCard6);
            this.tabPage1.Controls.Add(this.materialCard3);
            this.tabPage1.Controls.Add(this.materialCard2);
            this.tabPage1.Controls.Add(this.materialCard1);
            this.tabPage1.ImageKey = "home_house";
            this.tabPage1.Location = new System.Drawing.Point(4, 39);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1705, 983);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Home";
            // 
            // materialCard3
            // 
            this.materialCard3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard3.Controls.Add(this.card_datos);
            this.materialCard3.Controls.Add(this.materialCard5);
            this.materialCard3.Controls.Add(this.Card_hrs);
            this.materialCard3.Controls.Add(this.materialCard4);
            this.materialCard3.Controls.Add(this.lbl_meta);
            this.materialCard3.Controls.Add(this.cb_OP);
            this.materialCard3.Depth = 0;
            this.materialCard3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard3.Location = new System.Drawing.Point(14, 283);
            this.materialCard3.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard3.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard3.Name = "materialCard3";
            this.materialCard3.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard3.Size = new System.Drawing.Size(755, 559);
            this.materialCard3.TabIndex = 3;
            // 
            // pnl_principal
            // 
            this.pnl_principal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.pnl_principal.Controls.Add(this.txt_hr_Disp_TM);
            this.pnl_principal.Controls.Add(this.txt_meta_p);
            this.pnl_principal.Controls.Add(this.Mask_txt_hr1);
            this.pnl_principal.Controls.Add(this.Mask_txt_hr2);
            this.pnl_principal.Controls.Add(this.txt_hr_disponible);
            this.pnl_principal.Location = new System.Drawing.Point(0, 0);
            this.pnl_principal.Name = "pnl_principal";
            this.pnl_principal.Size = new System.Drawing.Size(369, 191);
            this.pnl_principal.TabIndex = 5;
            // 
            // txt_hr_Disp_TM
            // 
            this.txt_hr_Disp_TM.AnimateReadOnly = false;
            this.txt_hr_Disp_TM.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_hr_Disp_TM.Depth = 0;
            this.txt_hr_Disp_TM.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_hr_Disp_TM.Hint = "Horas disponibles menos tiempo muerto";
            this.txt_hr_Disp_TM.LeadingIcon = null;
            this.txt_hr_Disp_TM.Location = new System.Drawing.Point(17, 121);
            this.txt_hr_Disp_TM.MaxLength = 50;
            this.txt_hr_Disp_TM.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_hr_Disp_TM.Multiline = false;
            this.txt_hr_Disp_TM.Name = "txt_hr_Disp_TM";
            this.txt_hr_Disp_TM.Size = new System.Drawing.Size(335, 50);
            this.txt_hr_Disp_TM.TabIndex = 4;
            this.txt_hr_Disp_TM.Text = "";
            this.txt_hr_Disp_TM.TrailingIcon = null;
            // 
            // txt_meta_p
            // 
            this.txt_meta_p.AnimateReadOnly = false;
            this.txt_meta_p.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_meta_p.Depth = 0;
            this.txt_meta_p.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_meta_p.Hint = "Meta programada";
            this.txt_meta_p.LeadingIcon = null;
            this.txt_meta_p.Location = new System.Drawing.Point(195, 65);
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
            this.txt_hr_disponible.Location = new System.Drawing.Point(17, 65);
            this.txt_hr_disponible.MaxLength = 50;
            this.txt_hr_disponible.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_hr_disponible.Multiline = false;
            this.txt_hr_disponible.Name = "txt_hr_disponible";
            this.txt_hr_disponible.Size = new System.Drawing.Size(157, 50);
            this.txt_hr_disponible.TabIndex = 2;
            this.txt_hr_disponible.Text = "";
            this.txt_hr_disponible.TrailingIcon = null;
            // 
            // Card_hrs
            // 
            this.Card_hrs.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.Card_hrs.Controls.Add(this.pnl_principal);
            this.Card_hrs.Depth = 0;
            this.Card_hrs.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Card_hrs.Location = new System.Drawing.Point(367, 74);
            this.Card_hrs.Margin = new System.Windows.Forms.Padding(14);
            this.Card_hrs.MouseState = MaterialSkin.MouseState.HOVER;
            this.Card_hrs.Name = "Card_hrs";
            this.Card_hrs.Padding = new System.Windows.Forms.Padding(14);
            this.Card_hrs.Size = new System.Drawing.Size(369, 191);
            this.Card_hrs.TabIndex = 4;
            // 
            // materialCard4
            // 
            this.materialCard4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard4.Controls.Add(this.pnl_Metahora);
            this.materialCard4.Depth = 0;
            this.materialCard4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard4.Location = new System.Drawing.Point(631, 11);
            this.materialCard4.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard4.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard4.Name = "materialCard4";
            this.materialCard4.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard4.Size = new System.Drawing.Size(105, 46);
            this.materialCard4.TabIndex = 3;
            // 
            // pnl_Metahora
            // 
            this.pnl_Metahora.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.pnl_Metahora.Controls.Add(this.lbl_meta2);
            this.pnl_Metahora.Location = new System.Drawing.Point(0, 0);
            this.pnl_Metahora.Name = "pnl_Metahora";
            this.pnl_Metahora.Size = new System.Drawing.Size(105, 46);
            this.pnl_Metahora.TabIndex = 0;
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
            this.lbl_meta.Location = new System.Drawing.Point(482, 24);
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
            this.cb_OP.Location = new System.Drawing.Point(35, 11);
            this.cb_OP.MaxDropDownItems = 4;
            this.cb_OP.MouseState = MaterialSkin.MouseState.OUT;
            this.cb_OP.Name = "cb_OP";
            this.cb_OP.Size = new System.Drawing.Size(250, 49);
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
            this.materialCard2.Size = new System.Drawing.Size(755, 132);
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
            this.cb_Turno.Location = new System.Drawing.Point(49, 66);
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
            this.dtp1.Location = new System.Drawing.Point(384, 17);
            this.dtp1.Name = "dtp1";
            this.dtp1.Size = new System.Drawing.Size(335, 20);
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
            this.materialCard1.Size = new System.Drawing.Size(752, 78);
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
            // materialCard5
            // 
            this.materialCard5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard5.Controls.Add(this.pnl_ftt);
            this.materialCard5.Controls.Add(this.txt_Up_time);
            this.materialCard5.Controls.Add(this.txt_kg_F_Espec);
            this.materialCard5.Controls.Add(this.txt_kg_ok);
            this.materialCard5.Controls.Add(this.txt_perdida);
            this.materialCard5.Depth = 0;
            this.materialCard5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard5.Location = new System.Drawing.Point(18, 282);
            this.materialCard5.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard5.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard5.Name = "materialCard5";
            this.materialCard5.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard5.Size = new System.Drawing.Size(718, 200);
            this.materialCard5.TabIndex = 8;
            // 
            // card_datos
            // 
            this.card_datos.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.card_datos.Controls.Add(this.Txt_Personal_Op);
            this.card_datos.Controls.Add(this.txt_Produccion);
            this.card_datos.Controls.Add(this.Txt_Merma);
            this.card_datos.Depth = 0;
            this.card_datos.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.card_datos.Location = new System.Drawing.Point(18, 74);
            this.card_datos.Margin = new System.Windows.Forms.Padding(14);
            this.card_datos.MouseState = MaterialSkin.MouseState.HOVER;
            this.card_datos.Name = "card_datos";
            this.card_datos.Padding = new System.Windows.Forms.Padding(14);
            this.card_datos.Size = new System.Drawing.Size(321, 191);
            this.card_datos.TabIndex = 9;
            // 
            // txt_perdida
            // 
            this.txt_perdida.AnimateReadOnly = false;
            this.txt_perdida.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_perdida.Depth = 0;
            this.txt_perdida.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_perdida.Hint = "Perdida";
            this.txt_perdida.LeadingIcon = null;
            this.txt_perdida.Location = new System.Drawing.Point(17, 73);
            this.txt_perdida.MaxLength = 50;
            this.txt_perdida.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_perdida.Multiline = false;
            this.txt_perdida.Name = "txt_perdida";
            this.txt_perdida.Size = new System.Drawing.Size(234, 50);
            this.txt_perdida.TabIndex = 0;
            this.txt_perdida.Text = "";
            this.txt_perdida.TrailingIcon = null;
            // 
            // txt_ftt
            // 
            this.txt_ftt.AnimateReadOnly = false;
            this.txt_ftt.BackColor = System.Drawing.SystemColors.Info;
            this.txt_ftt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_ftt.Depth = 0;
            this.txt_ftt.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_ftt.Hint = "FTT";
            this.txt_ftt.LeadingIcon = null;
            this.txt_ftt.Location = new System.Drawing.Point(0, 2);
            this.txt_ftt.MaxLength = 50;
            this.txt_ftt.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_ftt.Multiline = false;
            this.txt_ftt.Name = "txt_ftt";
            this.txt_ftt.Size = new System.Drawing.Size(234, 50);
            this.txt_ftt.TabIndex = 1;
            this.txt_ftt.Text = "";
            this.txt_ftt.TrailingIcon = null;
            // 
            // txt_kg_ok
            // 
            this.txt_kg_ok.AnimateReadOnly = false;
            this.txt_kg_ok.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_kg_ok.Depth = 0;
            this.txt_kg_ok.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_kg_ok.Hint = "Kilogramos Ok";
            this.txt_kg_ok.LeadingIcon = null;
            this.txt_kg_ok.Location = new System.Drawing.Point(366, 17);
            this.txt_kg_ok.MaxLength = 50;
            this.txt_kg_ok.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_kg_ok.Multiline = false;
            this.txt_kg_ok.Name = "txt_kg_ok";
            this.txt_kg_ok.Size = new System.Drawing.Size(234, 50);
            this.txt_kg_ok.TabIndex = 2;
            this.txt_kg_ok.Text = "";
            this.txt_kg_ok.TrailingIcon = null;
            // 
            // txt_kg_F_Espec
            // 
            this.txt_kg_F_Espec.AnimateReadOnly = false;
            this.txt_kg_F_Espec.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_kg_F_Espec.Depth = 0;
            this.txt_kg_F_Espec.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_kg_F_Espec.Hint = "Kg Fuera de Especificación";
            this.txt_kg_F_Espec.LeadingIcon = null;
            this.txt_kg_F_Espec.Location = new System.Drawing.Point(366, 73);
            this.txt_kg_F_Espec.MaxLength = 50;
            this.txt_kg_F_Espec.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_kg_F_Espec.Multiline = false;
            this.txt_kg_F_Espec.Name = "txt_kg_F_Espec";
            this.txt_kg_F_Espec.Size = new System.Drawing.Size(234, 50);
            this.txt_kg_F_Espec.TabIndex = 3;
            this.txt_kg_F_Espec.Text = "";
            this.txt_kg_F_Espec.TrailingIcon = null;
            // 
            // txt_Up_time
            // 
            this.txt_Up_time.AnimateReadOnly = false;
            this.txt_Up_time.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_Up_time.Depth = 0;
            this.txt_Up_time.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_Up_time.Hint = "Up Time";
            this.txt_Up_time.LeadingIcon = null;
            this.txt_Up_time.Location = new System.Drawing.Point(17, 129);
            this.txt_Up_time.MaxLength = 50;
            this.txt_Up_time.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_Up_time.Multiline = false;
            this.txt_Up_time.Name = "txt_Up_time";
            this.txt_Up_time.Size = new System.Drawing.Size(234, 50);
            this.txt_Up_time.TabIndex = 4;
            this.txt_Up_time.Text = "";
            this.txt_Up_time.TrailingIcon = null;
            // 
            // txt_Produccion
            // 
            this.txt_Produccion.AnimateReadOnly = false;
            this.txt_Produccion.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_Produccion.Depth = 0;
            this.txt_Produccion.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_Produccion.Hint = "Resultado Producción";
            this.txt_Produccion.LeadingIcon = null;
            this.txt_Produccion.Location = new System.Drawing.Point(17, 9);
            this.txt_Produccion.MaxLength = 50;
            this.txt_Produccion.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_Produccion.Multiline = false;
            this.txt_Produccion.Name = "txt_Produccion";
            this.txt_Produccion.Size = new System.Drawing.Size(250, 50);
            this.txt_Produccion.TabIndex = 8;
            this.txt_Produccion.Text = "";
            this.txt_Produccion.TrailingIcon = null;
            // 
            // Txt_Merma
            // 
            this.Txt_Merma.AnimateReadOnly = false;
            this.Txt_Merma.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Txt_Merma.Depth = 0;
            this.Txt_Merma.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.Txt_Merma.Hint = "Merma";
            this.Txt_Merma.LeadingIcon = null;
            this.Txt_Merma.Location = new System.Drawing.Point(17, 65);
            this.Txt_Merma.MaxLength = 50;
            this.Txt_Merma.MouseState = MaterialSkin.MouseState.OUT;
            this.Txt_Merma.Multiline = false;
            this.Txt_Merma.Name = "Txt_Merma";
            this.Txt_Merma.Size = new System.Drawing.Size(250, 50);
            this.Txt_Merma.TabIndex = 9;
            this.Txt_Merma.Text = "";
            this.Txt_Merma.TrailingIcon = null;
            // 
            // Txt_Personal_Op
            // 
            this.Txt_Personal_Op.AnimateReadOnly = false;
            this.Txt_Personal_Op.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Txt_Personal_Op.Depth = 0;
            this.Txt_Personal_Op.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.Txt_Personal_Op.Hint = "Personal Operativo";
            this.Txt_Personal_Op.LeadingIcon = null;
            this.Txt_Personal_Op.Location = new System.Drawing.Point(17, 121);
            this.Txt_Personal_Op.MaxLength = 50;
            this.Txt_Personal_Op.MouseState = MaterialSkin.MouseState.OUT;
            this.Txt_Personal_Op.Multiline = false;
            this.Txt_Personal_Op.Name = "Txt_Personal_Op";
            this.Txt_Personal_Op.Size = new System.Drawing.Size(250, 50);
            this.Txt_Personal_Op.TabIndex = 10;
            this.Txt_Personal_Op.Text = "";
            this.Txt_Personal_Op.TrailingIcon = null;
            // 
            // pnl_ftt
            // 
            this.pnl_ftt.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.pnl_ftt.Controls.Add(this.txt_ftt);
            this.pnl_ftt.Location = new System.Drawing.Point(17, 17);
            this.pnl_ftt.Name = "pnl_ftt";
            this.pnl_ftt.Size = new System.Drawing.Size(234, 50);
            this.pnl_ftt.TabIndex = 5;
            // 
            // materialCard6
            // 
            this.materialCard6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard6.Controls.Add(this.txt_TM_calidad);
            this.materialCard6.Controls.Add(this.card_dgv_calidad);
            this.materialCard6.Controls.Add(this.txt_TM_operativo);
            this.materialCard6.Controls.Add(this.card_dgv_operativo);
            this.materialCard6.Controls.Add(this.txt_TM_almacen);
            this.materialCard6.Controls.Add(this.card_dgv_almacen);
            this.materialCard6.Controls.Add(this.lbl_min_comida);
            this.materialCard6.Controls.Add(this.txt_TM_mecanico);
            this.materialCard6.Controls.Add(this.card_dgv_mecanico);
            this.materialCard6.Controls.Add(this.txt_Tiempo_comida);
            this.materialCard6.Depth = 0;
            this.materialCard6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard6.Location = new System.Drawing.Point(797, 123);
            this.materialCard6.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard6.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard6.Name = "materialCard6";
            this.materialCard6.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard6.Size = new System.Drawing.Size(874, 719);
            this.materialCard6.TabIndex = 4;
            // 
            // txt_Tiempo_comida
            // 
            this.txt_Tiempo_comida.AnimateReadOnly = false;
            this.txt_Tiempo_comida.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_Tiempo_comida.Depth = 0;
            this.txt_Tiempo_comida.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_Tiempo_comida.Hint = "Tiempo muerto por comida";
            this.txt_Tiempo_comida.LeadingIcon = null;
            this.txt_Tiempo_comida.Location = new System.Drawing.Point(17, 46);
            this.txt_Tiempo_comida.MaxLength = 50;
            this.txt_Tiempo_comida.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_Tiempo_comida.Multiline = false;
            this.txt_Tiempo_comida.Name = "txt_Tiempo_comida";
            this.txt_Tiempo_comida.Size = new System.Drawing.Size(226, 50);
            this.txt_Tiempo_comida.TabIndex = 0;
            this.txt_Tiempo_comida.Text = "";
            this.txt_Tiempo_comida.TrailingIcon = null;
            this.txt_Tiempo_comida.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_Tiempo_comida_KeyPress);
            // 
            // dgv_mecanico
            // 
            this.dgv_mecanico.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_mecanico.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2});
            this.dgv_mecanico.Location = new System.Drawing.Point(0, 0);
            this.dgv_mecanico.Name = "dgv_mecanico";
            this.dgv_mecanico.Size = new System.Drawing.Size(386, 189);
            this.dgv_mecanico.TabIndex = 1;
            this.dgv_mecanico.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv_mecanico_EditingControlShowing);
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.HeaderText = "Minutos Detenidos";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column2.HeaderText = "Motivos Mecánicos";
            this.Column2.Name = "Column2";
            // 
            // card_dgv_mecanico
            // 
            this.card_dgv_mecanico.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.card_dgv_mecanico.Controls.Add(this.dgv_mecanico);
            this.card_dgv_mecanico.Depth = 0;
            this.card_dgv_mecanico.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.card_dgv_mecanico.Location = new System.Drawing.Point(17, 129);
            this.card_dgv_mecanico.Margin = new System.Windows.Forms.Padding(14);
            this.card_dgv_mecanico.MouseState = MaterialSkin.MouseState.HOVER;
            this.card_dgv_mecanico.Name = "card_dgv_mecanico";
            this.card_dgv_mecanico.Padding = new System.Windows.Forms.Padding(14);
            this.card_dgv_mecanico.Size = new System.Drawing.Size(386, 189);
            this.card_dgv_mecanico.TabIndex = 2;
            // 
            // txt_TM_mecanico
            // 
            this.txt_TM_mecanico.AnimateReadOnly = false;
            this.txt_TM_mecanico.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_TM_mecanico.Depth = 0;
            this.txt_TM_mecanico.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_TM_mecanico.Hint = "Minutos Totales";
            this.txt_TM_mecanico.LeadingIcon = null;
            this.txt_TM_mecanico.Location = new System.Drawing.Point(17, 324);
            this.txt_TM_mecanico.MaxLength = 50;
            this.txt_TM_mecanico.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_TM_mecanico.Multiline = false;
            this.txt_TM_mecanico.Name = "txt_TM_mecanico";
            this.txt_TM_mecanico.Size = new System.Drawing.Size(386, 50);
            this.txt_TM_mecanico.TabIndex = 3;
            this.txt_TM_mecanico.Text = "";
            this.txt_TM_mecanico.TrailingIcon = null;
            // 
            // lbl_min_comida
            // 
            this.lbl_min_comida.AutoSize = true;
            this.lbl_min_comida.Depth = 0;
            this.lbl_min_comida.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lbl_min_comida.Location = new System.Drawing.Point(249, 65);
            this.lbl_min_comida.MouseState = MaterialSkin.MouseState.HOVER;
            this.lbl_min_comida.Name = "lbl_min_comida";
            this.lbl_min_comida.Size = new System.Drawing.Size(59, 19);
            this.lbl_min_comida.TabIndex = 4;
            this.lbl_min_comida.Text = "Minutos";
            // 
            // card_dgv_almacen
            // 
            this.card_dgv_almacen.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.card_dgv_almacen.Controls.Add(this.dgv_almacen);
            this.card_dgv_almacen.Depth = 0;
            this.card_dgv_almacen.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.card_dgv_almacen.Location = new System.Drawing.Point(431, 129);
            this.card_dgv_almacen.Margin = new System.Windows.Forms.Padding(14);
            this.card_dgv_almacen.MouseState = MaterialSkin.MouseState.HOVER;
            this.card_dgv_almacen.Name = "card_dgv_almacen";
            this.card_dgv_almacen.Padding = new System.Windows.Forms.Padding(14);
            this.card_dgv_almacen.Size = new System.Drawing.Size(386, 189);
            this.card_dgv_almacen.TabIndex = 5;
            // 
            // dgv_almacen
            // 
            this.dgv_almacen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_almacen.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.dgv_almacen.Location = new System.Drawing.Point(0, 0);
            this.dgv_almacen.Name = "dgv_almacen";
            this.dgv_almacen.Size = new System.Drawing.Size(386, 189);
            this.dgv_almacen.TabIndex = 1;
            this.dgv_almacen.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv_almacen_EditingControlShowing);
            // 
            // txt_TM_almacen
            // 
            this.txt_TM_almacen.AnimateReadOnly = false;
            this.txt_TM_almacen.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_TM_almacen.Depth = 0;
            this.txt_TM_almacen.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_TM_almacen.Hint = "Minutos Totales";
            this.txt_TM_almacen.LeadingIcon = null;
            this.txt_TM_almacen.Location = new System.Drawing.Point(431, 324);
            this.txt_TM_almacen.MaxLength = 50;
            this.txt_TM_almacen.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_TM_almacen.Multiline = false;
            this.txt_TM_almacen.Name = "txt_TM_almacen";
            this.txt_TM_almacen.Size = new System.Drawing.Size(386, 50);
            this.txt_TM_almacen.TabIndex = 6;
            this.txt_TM_almacen.Text = "";
            this.txt_TM_almacen.TrailingIcon = null;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn1.HeaderText = "Minutos Detenidos";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn2.HeaderText = "Motivos Almacen";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // txt_TM_operativo
            // 
            this.txt_TM_operativo.AnimateReadOnly = false;
            this.txt_TM_operativo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_TM_operativo.Depth = 0;
            this.txt_TM_operativo.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_TM_operativo.Hint = "Minutos Totales";
            this.txt_TM_operativo.LeadingIcon = null;
            this.txt_TM_operativo.Location = new System.Drawing.Point(17, 617);
            this.txt_TM_operativo.MaxLength = 50;
            this.txt_TM_operativo.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_TM_operativo.Multiline = false;
            this.txt_TM_operativo.Name = "txt_TM_operativo";
            this.txt_TM_operativo.Size = new System.Drawing.Size(386, 50);
            this.txt_TM_operativo.TabIndex = 8;
            this.txt_TM_operativo.Text = "";
            this.txt_TM_operativo.TrailingIcon = null;
            // 
            // card_dgv_operativo
            // 
            this.card_dgv_operativo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.card_dgv_operativo.Controls.Add(this.dgv_operativo);
            this.card_dgv_operativo.Depth = 0;
            this.card_dgv_operativo.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.card_dgv_operativo.Location = new System.Drawing.Point(17, 422);
            this.card_dgv_operativo.Margin = new System.Windows.Forms.Padding(14);
            this.card_dgv_operativo.MouseState = MaterialSkin.MouseState.HOVER;
            this.card_dgv_operativo.Name = "card_dgv_operativo";
            this.card_dgv_operativo.Padding = new System.Windows.Forms.Padding(14);
            this.card_dgv_operativo.Size = new System.Drawing.Size(386, 189);
            this.card_dgv_operativo.TabIndex = 7;
            // 
            // dgv_operativo
            // 
            this.dgv_operativo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_operativo.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4});
            this.dgv_operativo.Location = new System.Drawing.Point(0, 0);
            this.dgv_operativo.Name = "dgv_operativo";
            this.dgv_operativo.Size = new System.Drawing.Size(386, 189);
            this.dgv_operativo.TabIndex = 1;
            this.dgv_operativo.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv_operativo_EditingControlShowing);
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn3.HeaderText = "Minutos Detenidos";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn4.HeaderText = "Motivos Operativos";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // txt_TM_calidad
            // 
            this.txt_TM_calidad.AnimateReadOnly = false;
            this.txt_TM_calidad.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt_TM_calidad.Depth = 0;
            this.txt_TM_calidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.txt_TM_calidad.Hint = "Minutos Totales";
            this.txt_TM_calidad.LeadingIcon = null;
            this.txt_TM_calidad.Location = new System.Drawing.Point(431, 617);
            this.txt_TM_calidad.MaxLength = 50;
            this.txt_TM_calidad.MouseState = MaterialSkin.MouseState.OUT;
            this.txt_TM_calidad.Multiline = false;
            this.txt_TM_calidad.Name = "txt_TM_calidad";
            this.txt_TM_calidad.Size = new System.Drawing.Size(386, 50);
            this.txt_TM_calidad.TabIndex = 10;
            this.txt_TM_calidad.Text = "";
            this.txt_TM_calidad.TrailingIcon = null;
            // 
            // card_dgv_calidad
            // 
            this.card_dgv_calidad.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.card_dgv_calidad.Controls.Add(this.dgv_calidad);
            this.card_dgv_calidad.Depth = 0;
            this.card_dgv_calidad.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.card_dgv_calidad.Location = new System.Drawing.Point(431, 422);
            this.card_dgv_calidad.Margin = new System.Windows.Forms.Padding(14);
            this.card_dgv_calidad.MouseState = MaterialSkin.MouseState.HOVER;
            this.card_dgv_calidad.Name = "card_dgv_calidad";
            this.card_dgv_calidad.Padding = new System.Windows.Forms.Padding(14);
            this.card_dgv_calidad.Size = new System.Drawing.Size(386, 189);
            this.card_dgv_calidad.TabIndex = 9;
            // 
            // dgv_calidad
            // 
            this.dgv_calidad.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_calidad.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6});
            this.dgv_calidad.Location = new System.Drawing.Point(0, 0);
            this.dgv_calidad.Name = "dgv_calidad";
            this.dgv_calidad.Size = new System.Drawing.Size(386, 189);
            this.dgv_calidad.TabIndex = 1;
            this.dgv_calidad.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv_calidad_EditingControlShowing);
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn5.HeaderText = "Minutos Detenidos";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn6.HeaderText = "Motivos Calidad";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            // 
            // Form_principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1719, 1093);
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
            this.Card_hrs.ResumeLayout(false);
            this.materialCard4.ResumeLayout(false);
            this.pnl_Metahora.ResumeLayout(false);
            this.pnl_Metahora.PerformLayout();
            this.materialCard2.ResumeLayout(false);
            this.materialCard1.ResumeLayout(false);
            this.materialCard1.PerformLayout();
            this.materialCard5.ResumeLayout(false);
            this.card_datos.ResumeLayout(false);
            this.pnl_ftt.ResumeLayout(false);
            this.materialCard6.ResumeLayout(false);
            this.materialCard6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_mecanico)).EndInit();
            this.card_dgv_mecanico.ResumeLayout(false);
            this.card_dgv_almacen.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_almacen)).EndInit();
            this.card_dgv_operativo.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_operativo)).EndInit();
            this.card_dgv_calidad.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_calidad)).EndInit();
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
        private MaterialSkin.Controls.MaterialCard Card_hrs;
        private System.Windows.Forms.Panel pnl_principal;
        private MaterialSkin.Controls.MaterialTextBox txt_hr_Disp_TM;
        private System.Windows.Forms.Panel pnl_Metahora;
        private MaterialSkin.Controls.MaterialCard card_datos;
        private MaterialSkin.Controls.MaterialCard materialCard5;
        private MaterialSkin.Controls.MaterialTextBox txt_kg_F_Espec;
        private MaterialSkin.Controls.MaterialTextBox txt_kg_ok;
        private MaterialSkin.Controls.MaterialTextBox txt_ftt;
        private MaterialSkin.Controls.MaterialTextBox txt_perdida;
        private MaterialSkin.Controls.MaterialTextBox txt_Up_time;
        private MaterialSkin.Controls.MaterialTextBox txt_Produccion;
        private MaterialSkin.Controls.MaterialTextBox Txt_Personal_Op;
        private MaterialSkin.Controls.MaterialTextBox Txt_Merma;
        private System.Windows.Forms.Panel pnl_ftt;
        private MaterialSkin.Controls.MaterialCard materialCard6;
        private MaterialSkin.Controls.MaterialTextBox txt_Tiempo_comida;
        private System.Windows.Forms.DataGridView dgv_mecanico;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private MaterialSkin.Controls.MaterialCard card_dgv_mecanico;
        private MaterialSkin.Controls.MaterialLabel lbl_min_comida;
        private MaterialSkin.Controls.MaterialTextBox txt_TM_mecanico;
        private MaterialSkin.Controls.MaterialCard card_dgv_almacen;
        private System.Windows.Forms.DataGridView dgv_almacen;
        private MaterialSkin.Controls.MaterialTextBox txt_TM_almacen;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private MaterialSkin.Controls.MaterialTextBox txt_TM_operativo;
        private MaterialSkin.Controls.MaterialCard card_dgv_operativo;
        private System.Windows.Forms.DataGridView dgv_operativo;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private MaterialSkin.Controls.MaterialTextBox txt_TM_calidad;
        private MaterialSkin.Controls.MaterialCard card_dgv_calidad;
        private System.Windows.Forms.DataGridView dgv_calidad;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
    }
}

