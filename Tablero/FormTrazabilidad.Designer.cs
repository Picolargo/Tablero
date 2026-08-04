namespace Tablero
{
    partial class FormTrazabilidad
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormTrazabilidad));
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition2 = new Telerik.WinControls.UI.TableViewDefinition();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnLimpiar = new MaterialSkin.Controls.MaterialButton();
            this.btnBuscar = new MaterialSkin.Controls.MaterialButton();
            this.cmbUsuario = new MaterialSkin.Controls.MaterialComboBox();
            this.lblUsuario = new MaterialSkin.Controls.MaterialLabel();
            this.dtpHasta = new MetroFramework.Controls.MetroDateTime();
            this.lblHasta = new MaterialSkin.Controls.MaterialLabel();
            this.dtpDesde = new MetroFramework.Controls.MetroDateTime();
            this.lblDesde = new MaterialSkin.Controls.MaterialLabel();
            this.materialCardResultados = new MaterialSkin.Controls.MaterialCard();
            this.btnExportExcel = new MaterialSkin.Controls.MaterialButton();
            this.lblTotalRegistros = new MaterialSkin.Controls.MaterialLabel();
            this.rgvTrazabilidad = new Telerik.WinControls.UI.RadGridView();
            this.materialCardFiltros = new MaterialSkin.Controls.MaterialCard();
            this.materialCard1 = new MaterialSkin.Controls.MaterialCard();
            this.materialCardResultados.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rgvTrazabilidad)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rgvTrazabilidad.MasterTemplate)).BeginInit();
            this.materialCardFiltros.SuspendLayout();
            this.materialCard1.SuspendLayout();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "statistics_251075");
            this.imageList1.Images.SetKeyName(1, "search_18113826");
            // 
            // btnLimpiar
            // 
            this.btnLimpiar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnLimpiar.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.btnLimpiar.Depth = 0;
            this.btnLimpiar.Enabled = false;
            this.btnLimpiar.HighEmphasis = true;
            this.btnLimpiar.Icon = global::Tablero.Properties.Resources.filter_remove;
            this.btnLimpiar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnLimpiar.Location = new System.Drawing.Point(436, 142);
            this.btnLimpiar.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.btnLimpiar.MouseState = MaterialSkin.MouseState.HOVER;
            this.btnLimpiar.Name = "btnLimpiar";
            this.btnLimpiar.NoAccentTextColor = System.Drawing.Color.Empty;
            this.btnLimpiar.Size = new System.Drawing.Size(107, 36);
            this.btnLimpiar.TabIndex = 9;
            this.btnLimpiar.Text = "Limpiar";
            this.btnLimpiar.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.btnLimpiar.UseAccentColor = false;
            this.btnLimpiar.UseVisualStyleBackColor = true;
            this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);
            // 
            // btnBuscar
            // 
            this.btnBuscar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnBuscar.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.btnBuscar.Depth = 0;
            this.btnBuscar.HighEmphasis = true;
            this.btnBuscar.Icon = global::Tablero.Properties.Resources._8666693_search_icon;
            this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnBuscar.Location = new System.Drawing.Point(316, 142);
            this.btnBuscar.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.btnBuscar.MouseState = MaterialSkin.MouseState.HOVER;
            this.btnBuscar.Name = "btnBuscar";
            this.btnBuscar.NoAccentTextColor = System.Drawing.Color.Empty;
            this.btnBuscar.Size = new System.Drawing.Size(105, 36);
            this.btnBuscar.TabIndex = 8;
            this.btnBuscar.Text = "Buscar";
            this.btnBuscar.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.btnBuscar.UseAccentColor = false;
            this.btnBuscar.UseVisualStyleBackColor = true;
            this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
            // 
            // cmbUsuario
            // 
            this.cmbUsuario.AutoResize = false;
            this.cmbUsuario.BackColor = System.Drawing.Color.White;
            this.cmbUsuario.Depth = 0;
            this.cmbUsuario.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.cmbUsuario.DropDownHeight = 174;
            this.cmbUsuario.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbUsuario.DropDownWidth = 121;
            this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            this.cmbUsuario.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.cmbUsuario.FormattingEnabled = true;
            this.cmbUsuario.IntegralHeight = false;
            this.cmbUsuario.ItemHeight = 43;
            this.cmbUsuario.Location = new System.Drawing.Point(26, 129);
            this.cmbUsuario.MaxDropDownItems = 4;
            this.cmbUsuario.MouseState = MaterialSkin.MouseState.OUT;
            this.cmbUsuario.Name = "cmbUsuario";
            this.cmbUsuario.Size = new System.Drawing.Size(250, 49);
            this.cmbUsuario.StartIndex = -1;
            this.cmbUsuario.TabIndex = 7;
            // 
            // lblUsuario
            // 
            this.lblUsuario.AutoSize = true;
            this.lblUsuario.Depth = 0;
            this.lblUsuario.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lblUsuario.Location = new System.Drawing.Point(26, 104);
            this.lblUsuario.MouseState = MaterialSkin.MouseState.HOVER;
            this.lblUsuario.Name = "lblUsuario";
            this.lblUsuario.Size = new System.Drawing.Size(55, 19);
            this.lblUsuario.TabIndex = 6;
            this.lblUsuario.Text = "Usuario";
            // 
            // dtpHasta
            // 
            this.dtpHasta.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold);
            this.dtpHasta.FontSize = MetroFramework.MetroDateTimeSize.Tall;
            this.dtpHasta.FontWeight = MetroFramework.MetroDateTimeWeight.Light;
            this.dtpHasta.Location = new System.Drawing.Point(316, 39);
            this.dtpHasta.MinimumSize = new System.Drawing.Size(0, 35);
            this.dtpHasta.Name = "dtpHasta";
            this.dtpHasta.Size = new System.Drawing.Size(250, 35);
            this.dtpHasta.TabIndex = 3;
            // 
            // lblHasta
            // 
            this.lblHasta.AutoSize = true;
            this.lblHasta.Depth = 0;
            this.lblHasta.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lblHasta.Location = new System.Drawing.Point(316, 14);
            this.lblHasta.MouseState = MaterialSkin.MouseState.HOVER;
            this.lblHasta.Name = "lblHasta";
            this.lblHasta.Size = new System.Drawing.Size(43, 19);
            this.lblHasta.TabIndex = 2;
            this.lblHasta.Text = "Hasta";
            // 
            // dtpDesde
            // 
            this.dtpDesde.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold);
            this.dtpDesde.FontSize = MetroFramework.MetroDateTimeSize.Tall;
            this.dtpDesde.FontWeight = MetroFramework.MetroDateTimeWeight.Light;
            this.dtpDesde.Location = new System.Drawing.Point(26, 39);
            this.dtpDesde.MinimumSize = new System.Drawing.Size(0, 35);
            this.dtpDesde.Name = "dtpDesde";
            this.dtpDesde.Size = new System.Drawing.Size(250, 35);
            this.dtpDesde.TabIndex = 1;
            // 
            // lblDesde
            // 
            this.lblDesde.AutoSize = true;
            this.lblDesde.Depth = 0;
            this.lblDesde.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lblDesde.Location = new System.Drawing.Point(26, 14);
            this.lblDesde.MouseState = MaterialSkin.MouseState.HOVER;
            this.lblDesde.Name = "lblDesde";
            this.lblDesde.Size = new System.Drawing.Size(45, 19);
            this.lblDesde.TabIndex = 0;
            this.lblDesde.Text = "Desde";
            // 
            // materialCardResultados
            // 
            this.materialCardResultados.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCardResultados.Controls.Add(this.rgvTrazabilidad);
            this.materialCardResultados.Controls.Add(this.lblTotalRegistros);
            this.materialCardResultados.Controls.Add(this.btnExportExcel);
            this.materialCardResultados.Depth = 0;
            this.materialCardResultados.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.materialCardResultados.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCardResultados.Location = new System.Drawing.Point(14, 249);
            this.materialCardResultados.Margin = new System.Windows.Forms.Padding(14);
            this.materialCardResultados.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCardResultados.Name = "materialCardResultados";
            this.materialCardResultados.Padding = new System.Windows.Forms.Padding(14);
            this.materialCardResultados.Size = new System.Drawing.Size(1166, 420);
            this.materialCardResultados.TabIndex = 10;
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportExcel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnExportExcel.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.btnExportExcel.Depth = 0;
            this.btnExportExcel.Enabled = false;
            this.btnExportExcel.HighEmphasis = true;
            this.btnExportExcel.Icon = global::Tablero.Properties.Resources.excel_icon;
            this.btnExportExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportExcel.Location = new System.Drawing.Point(917, 13);
            this.btnExportExcel.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.btnExportExcel.MouseState = MaterialSkin.MouseState.HOVER;
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.NoAccentTextColor = System.Drawing.Color.Empty;
            this.btnExportExcel.Size = new System.Drawing.Size(184, 36);
            this.btnExportExcel.TabIndex = 11;
            this.btnExportExcel.Text = "Exportar a Excel";
            this.btnExportExcel.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.btnExportExcel.UseAccentColor = false;
            this.btnExportExcel.UseVisualStyleBackColor = true;
            // 
            // lblTotalRegistros
            // 
            this.lblTotalRegistros.AutoSize = true;
            this.lblTotalRegistros.Depth = 0;
            this.lblTotalRegistros.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lblTotalRegistros.Location = new System.Drawing.Point(14, 30);
            this.lblTotalRegistros.MouseState = MaterialSkin.MouseState.HOVER;
            this.lblTotalRegistros.Name = "lblTotalRegistros";
            this.lblTotalRegistros.Size = new System.Drawing.Size(120, 19);
            this.lblTotalRegistros.TabIndex = 10;
            this.lblTotalRegistros.Text = "Total: 0 registros";
            // 
            // rgvTrazabilidad
            // 
            this.rgvTrazabilidad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rgvTrazabilidad.Location = new System.Drawing.Point(14, 60);
            // 
            // 
            // 
            this.rgvTrazabilidad.MasterTemplate.ViewDefinition = tableViewDefinition2;
            this.rgvTrazabilidad.Name = "rgvTrazabilidad";
            this.rgvTrazabilidad.Size = new System.Drawing.Size(1135, 346);
            this.rgvTrazabilidad.TabIndex = 0;
            // 
            // materialCardFiltros
            // 
            this.materialCardFiltros.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCardFiltros.Controls.Add(this.materialCard1);
            this.materialCardFiltros.Controls.Add(this.materialCardResultados);
            this.materialCardFiltros.Depth = 0;
            this.materialCardFiltros.Dock = System.Windows.Forms.DockStyle.Fill;
            this.materialCardFiltros.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCardFiltros.Location = new System.Drawing.Point(3, 64);
            this.materialCardFiltros.Margin = new System.Windows.Forms.Padding(14);
            this.materialCardFiltros.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCardFiltros.Name = "materialCardFiltros";
            this.materialCardFiltros.Padding = new System.Windows.Forms.Padding(14);
            this.materialCardFiltros.Size = new System.Drawing.Size(1194, 683);
            this.materialCardFiltros.TabIndex = 0;
            // 
            // materialCard1
            // 
            this.materialCard1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.materialCard1.Controls.Add(this.lblDesde);
            this.materialCard1.Controls.Add(this.btnLimpiar);
            this.materialCard1.Controls.Add(this.btnBuscar);
            this.materialCard1.Controls.Add(this.dtpDesde);
            this.materialCard1.Controls.Add(this.cmbUsuario);
            this.materialCard1.Controls.Add(this.lblHasta);
            this.materialCard1.Controls.Add(this.lblUsuario);
            this.materialCard1.Controls.Add(this.dtpHasta);
            this.materialCard1.Depth = 0;
            this.materialCard1.Dock = System.Windows.Forms.DockStyle.Top;
            this.materialCard1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.materialCard1.Location = new System.Drawing.Point(14, 14);
            this.materialCard1.Margin = new System.Windows.Forms.Padding(14);
            this.materialCard1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialCard1.Name = "materialCard1";
            this.materialCard1.Padding = new System.Windows.Forms.Padding(14);
            this.materialCard1.Size = new System.Drawing.Size(1166, 207);
            this.materialCard1.TabIndex = 11;
            // 
            // FormTrazabilidad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 750);
            this.Controls.Add(this.materialCardFiltros);
            this.DrawerShowIconsWhenHidden = true;
            this.Name = "FormTrazabilidad";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Trazabilidad de Ediciones de Fichas";
            this.Load += new System.EventHandler(this.FormTrazabilidad_Load);
            this.materialCardResultados.ResumeLayout(false);
            this.materialCardResultados.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rgvTrazabilidad.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rgvTrazabilidad)).EndInit();
            this.materialCardFiltros.ResumeLayout(false);
            this.materialCard1.ResumeLayout(false);
            this.materialCard1.PerformLayout();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.ImageList imageList1;
        private MaterialSkin.Controls.MaterialCard materialCardFiltros;
        private MaterialSkin.Controls.MaterialCard materialCardResultados;
        private Telerik.WinControls.UI.RadGridView rgvTrazabilidad;
        private MaterialSkin.Controls.MaterialLabel lblTotalRegistros;
        private MaterialSkin.Controls.MaterialButton btnExportExcel;
        private MaterialSkin.Controls.MaterialLabel lblDesde;
        private MetroFramework.Controls.MetroDateTime dtpDesde;
        private MaterialSkin.Controls.MaterialLabel lblHasta;
        private MetroFramework.Controls.MetroDateTime dtpHasta;
        private MaterialSkin.Controls.MaterialLabel lblUsuario;
        private MaterialSkin.Controls.MaterialComboBox cmbUsuario;
        private MaterialSkin.Controls.MaterialButton btnBuscar;
        private MaterialSkin.Controls.MaterialButton btnLimpiar;
        private MaterialSkin.Controls.MaterialCard materialCard1;
    }
}