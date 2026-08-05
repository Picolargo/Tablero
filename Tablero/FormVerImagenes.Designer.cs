namespace Tablero
{
    partial class FormVerImagenes
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
            this.pictureBoxImagen = new System.Windows.Forms.PictureBox();
            this.btnAnterior = new System.Windows.Forms.Button();
            this.btnSiguiente = new System.Windows.Forms.Button();
            this.lblInfo = new System.Windows.Forms.Label();
            this.lblNoImagenes = new System.Windows.Forms.Label();
            this.listBoxMiniaturas = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxImagen)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();

            // pictureBoxImagen
            this.pictureBoxImagen.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            this.pictureBoxImagen.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBoxImagen.Location = new System.Drawing.Point(0, 0);
            this.pictureBoxImagen.Name = "pictureBoxImagen";
            this.pictureBoxImagen.Size = new System.Drawing.Size(800, 500);
            this.pictureBoxImagen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxImagen.TabIndex = 0;
            this.pictureBoxImagen.TabStop = false;
            this.pictureBoxImagen.Click += new System.EventHandler(this.pictureBoxImagen_Click);

            // btnAnterior
            this.btnAnterior.Enabled = false;
            this.btnAnterior.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.btnAnterior.Location = new System.Drawing.Point(10, 210);
            this.btnAnterior.Name = "btnAnterior";
            this.btnAnterior.Size = new System.Drawing.Size(60, 80);
            this.btnAnterior.TabIndex = 1;
            this.btnAnterior.Text = "◄";
            this.btnAnterior.UseVisualStyleBackColor = true;
            this.btnAnterior.Click += new System.EventHandler(this.btnAnterior_Click);

            // btnSiguiente
            this.btnSiguiente.Enabled = false;
            this.btnSiguiente.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.btnSiguiente.Location = new System.Drawing.Point(730, 210);
            this.btnSiguiente.Name = "btnSiguiente";
            this.btnSiguiente.Size = new System.Drawing.Size(60, 80);
            this.btnSiguiente.TabIndex = 2;
            this.btnSiguiente.Text = "►";
            this.btnSiguiente.UseVisualStyleBackColor = true;
            this.btnSiguiente.Click += new System.EventHandler(this.btnSiguiente_Click);

            // lblInfo
            this.lblInfo.BackColor = System.Drawing.Color.FromArgb(50, 50, 50);
            this.lblInfo.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.lblInfo.ForeColor = System.Drawing.Color.White;
            this.lblInfo.Location = new System.Drawing.Point(0, 500);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Padding = new System.Windows.Forms.Padding(10, 5, 0, 5);
            this.lblInfo.Size = new System.Drawing.Size(800, 30);
            this.lblInfo.TabIndex = 3;
            this.lblInfo.Text = "0/0";
            this.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            // lblNoImagenes
            this.lblNoImagenes.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            this.lblNoImagenes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblNoImagenes.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold);
            this.lblNoImagenes.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100);
            this.lblNoImagenes.Location = new System.Drawing.Point(0, 0);
            this.lblNoImagenes.Name = "lblNoImagenes";
            this.lblNoImagenes.Size = new System.Drawing.Size(800, 500);
            this.lblNoImagenes.TabIndex = 4;
            this.lblNoImagenes.Text = "No hay imágenes para esta ficha";
            this.lblNoImagenes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblNoImagenes.Visible = false;

            // listBoxMiniaturas
            this.listBoxMiniaturas.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxMiniaturas.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.listBoxMiniaturas.FormattingEnabled = true;
            this.listBoxMiniaturas.ItemHeight = 20;
            this.listBoxMiniaturas.Location = new System.Drawing.Point(0, 0);
            this.listBoxMiniaturas.Name = "listBoxMiniaturas";
            this.listBoxMiniaturas.Size = new System.Drawing.Size(250, 530);
            this.listBoxMiniaturas.TabIndex = 5;
            this.listBoxMiniaturas.SelectedIndexChanged += new System.EventHandler(this.listBoxMiniaturas_SelectedIndexChanged);

            // panel1
            this.panel1.Controls.Add(this.lblNoImagenes);
            this.panel1.Controls.Add(this.pictureBoxImagen);
            this.panel1.Controls.Add(this.btnAnterior);
            this.panel1.Controls.Add(this.btnSiguiente);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(260, 0);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(0, 0, 0, 30);
            this.panel1.Size = new System.Drawing.Size(800, 530);
            this.panel1.TabIndex = 6;

            // panel2
            this.panel2.Controls.Add(this.lblInfo);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(260, 530);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(800, 30);
            this.panel2.TabIndex = 7;

            // panel3
            this.panel3.Controls.Add(this.listBoxMiniaturas);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(260, 560);
            this.panel3.TabIndex = 8;

            // FormVerImagenes
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1060, 560);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(500, 400);
            this.Name = "FormVerImagenes";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Ver Imágenes";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormVerImagenes_FormClosing);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormVerImagenes_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxImagen)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBoxImagen;
        private System.Windows.Forms.Button btnAnterior;
        private System.Windows.Forms.Button btnSiguiente;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Label lblNoImagenes;
        private System.Windows.Forms.ListBox listBoxMiniaturas;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;

        /// <summary>
        /// Maneja las teclas de navegación
        /// </summary>
        private void FormVerImagenes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Right && btnSiguiente.Enabled)
                btnSiguiente_Click(null, null);
            else if (e.KeyCode == System.Windows.Forms.Keys.Left && btnAnterior.Enabled)
                btnAnterior_Click(null, null);
            else if (e.KeyCode == System.Windows.Forms.Keys.Escape)
                this.Close();
        }
    }
}