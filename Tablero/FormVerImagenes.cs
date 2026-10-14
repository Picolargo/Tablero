using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace Tablero
{
    public partial class FormVerImagenes : Form
    {
        private int _idFicha;
        private DataTable _imagenes;
        private string _connectionString;
        private int _imagenActualIndex = 0;
        private List<string> _rutasImagenes = new List<string>();
        private Form _formImagenAmpliada = null;
        private bool _imagenAmpliadaVisible = false;
        // ═══════════════════════════════════════════════════════════
        // VARIABLES PARA ZOOM Y ROTACIÓN
        // ═══════════════════════════════════════════════════════════
        private float _zoomAmpliado = 1.0f;
        private float _rotacionAmpliado = 0f;
        private Image _imagenOriginalAmpliada = null;
        private PictureBox _pbAmpliada = null;
        private Panel _panelScrollAmpliada = null;
        private Label _lblZoomInfo = null;
        // ═══════════════════════════════════════════════════════════

        public FormVerImagenes(int idFicha, string connectionString)
        {
            InitializeComponent();
            _idFicha = idFicha;
            _connectionString = connectionString;
            this.Text = $"Imágenes de Ficha #{idFicha}";
            CargarImagenes();
        }

        private void CargarImagenes()
        {
            try
            {
                DatabaseHelper dbHelper = new DatabaseHelper(_connectionString);
                _imagenes = dbHelper.ObtenerImagenesPorFicha(_idFicha);

                if (_imagenes == null || _imagenes.Rows.Count == 0)
                {
                    lblNoImagenes.Visible = true;
                    btnSiguiente.Enabled = false;
                    btnAnterior.Enabled = false;
                    pictureBoxImagen.Image = null;
                    lblInfo.Text = "0/0";
                    return;
                }

                lblNoImagenes.Visible = false;
                _rutasImagenes.Clear();

                foreach (DataRow row in _imagenes.Rows)
                {
                    string ruta = row["Ruta_Completa"].ToString();
                    if (File.Exists(ruta))
                    {
                        _rutasImagenes.Add(ruta);
                    }
                }

                if (_rutasImagenes.Count == 0)
                {
                    lblNoImagenes.Visible = true;
                    lblNoImagenes.Text = "Las imágenes no se encuentran en el servidor";
                    btnSiguiente.Enabled = false;
                    btnAnterior.Enabled = false;
                    pictureBoxImagen.Image = null;
                    lblInfo.Text = "0/0";
                    return;
                }

                _imagenActualIndex = 0;
                MostrarImagenActual();
                btnSiguiente.Enabled = _rutasImagenes.Count > 1;
                btnAnterior.Enabled = false;

                // Cargar lista de imágenes en el ListBox
                listBoxMiniaturas.Items.Clear();
                for (int i = 0; i < _rutasImagenes.Count; i++)
                {
                    string nombre = Path.GetFileName(_rutasImagenes[i]);
                    listBoxMiniaturas.Items.Add($"{i + 1}. {nombre}");
                }
                if (listBoxMiniaturas.Items.Count > 0)
                    listBoxMiniaturas.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar imágenes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MostrarImagenActual()
        {
            if (_rutasImagenes.Count == 0 || _imagenActualIndex < 0 || _imagenActualIndex >= _rutasImagenes.Count)
            {
                pictureBoxImagen.Image = null;
                lblInfo.Text = "0/0";
                return;
            }

            try
            {
                string ruta = _rutasImagenes[_imagenActualIndex];
                using (var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read))
                {
                    pictureBoxImagen.Image = Image.FromStream(stream);
                }
                pictureBoxImagen.SizeMode = PictureBoxSizeMode.Zoom;

                // Actualizar información
                FileInfo info = new FileInfo(ruta);
                lblInfo.Text = $"{_imagenActualIndex + 1}/{_rutasImagenes.Count} - {info.Length / 1024} KB";

                // Actualizar selección en el ListBox
                if (listBoxMiniaturas.SelectedIndex != _imagenActualIndex)
                    listBoxMiniaturas.SelectedIndex = _imagenActualIndex;
            }
            catch (Exception ex)
            {
                pictureBoxImagen.Image = null;
                lblInfo.Text = $"Error: {ex.Message}";
            }
        }

        private void btnAnterior_Click(object sender, EventArgs e)
        {
            if (_imagenActualIndex > 0)
            {
                _imagenActualIndex--;
                MostrarImagenActual();
                btnSiguiente.Enabled = true;
                btnAnterior.Enabled = _imagenActualIndex > 0;
            }
        }

        private void btnSiguiente_Click(object sender, EventArgs e)
        {
            if (_imagenActualIndex < _rutasImagenes.Count - 1)
            {
                _imagenActualIndex++;
                MostrarImagenActual();
                btnAnterior.Enabled = true;
                btnSiguiente.Enabled = _imagenActualIndex < _rutasImagenes.Count - 1;
            }
        }

        private void pictureBoxImagen_Click(object sender, EventArgs e)
        {
            // Ampliar imagen a pantalla completa
            if (pictureBoxImagen.Image != null)
            {
                MostrarImagenAmpliada();
            }
        }

        private void MostrarImagenAmpliada()
        {
            if (pictureBoxImagen.Image == null) return;
            if (_imagenAmpliadaVisible) { CerrarImagenAmpliada(); return; }

            // ═══ Resetear estado ═══
            _zoomAmpliado = 1.0f;
            _rotacionAmpliado = 0f;

            // COPIA INDEPENDIENTE de la imagen
            _imagenOriginalAmpliada = new Bitmap(pictureBoxImagen.Image);

            _formImagenAmpliada = new Form();
            _formImagenAmpliada.FormBorderStyle = FormBorderStyle.None;
            _formImagenAmpliada.BackColor = Color.Black;
            _formImagenAmpliada.StartPosition = FormStartPosition.CenterScreen;
            _formImagenAmpliada.TopMost = true;
            _formImagenAmpliada.KeyPreview = true;

            // ═══ Panel con scroll que contendrá la imagen ═══
            _panelScrollAmpliada = new Panel();
            _panelScrollAmpliada.Dock = DockStyle.Fill;
            _panelScrollAmpliada.AutoScroll = true;
            _panelScrollAmpliada.BackColor = Color.Black;

            // ═══ PictureBox con SizeMode.AutoSize ═══
            _pbAmpliada = new PictureBox();
            _pbAmpliada.SizeMode = PictureBoxSizeMode.AutoSize;
            _pbAmpliada.BackColor = Color.Black;
            _pbAmpliada.Location = new Point(0, 0);

            _panelScrollAmpliada.Controls.Add(_pbAmpliada);

            // ═══ Panel superior con controles ═══
            Panel panelControles = new Panel();
            panelControles.Dock = DockStyle.Top;
            panelControles.Height = 50;
            panelControles.BackColor = Color.FromArgb(40, 40, 40);

            Button btnRotarIzq = CrearBotonControl("↺  Rotar -90°", 20, 8);
            btnRotarIzq.Click += (s, ev) => { _rotacionAmpliado -= 90f; AplicarTransformacionAmpliada(); };

            Button btnRotarDer = CrearBotonControl("↻  Rotar +90°", 220, 8);
            btnRotarDer.Click += (s, ev) => { _rotacionAmpliado += 90f; AplicarTransformacionAmpliada(); };

            Button btnZoomIn = CrearBotonControl("＋  Zoom +", 420, 8);
            btnZoomIn.Click += (s, ev) => { _zoomAmpliado = Math.Min(_zoomAmpliado + 0.25f, 5.0f); AplicarTransformacionAmpliada(); };

            Button btnZoomOut = CrearBotonControl("－  Zoom -", 620, 8);
            btnZoomOut.Click += (s, ev) => { _zoomAmpliado = Math.Max(_zoomAmpliado - 0.25f, 0.25f); AplicarTransformacionAmpliada(); };

            Button btnReset = CrearBotonControl("⟲  Reset", 820, 8);
            btnReset.Click += (s, ev) =>
            {
                _zoomAmpliado = 1.0f;
                _rotacionAmpliado = 0f;
                AplicarTransformacionAmpliada();
            };

            // ═══════════════════════════════════════════════════════════
            // 🔥 NUEVO BOTÓN CERRAR
            // ═══════════════════════════════════════════════════════════
            Button btnCerrar = CrearBotonControl("✖  Cerrar", 1020, 8);
            btnCerrar.BackColor = Color.FromArgb(200, 50, 50);      // Rojo
            btnCerrar.FlatAppearance.BorderColor = Color.FromArgb(255, 80, 80);
            btnCerrar.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 70, 70);
            btnCerrar.Click += (s, ev) => CerrarImagenAmpliada();
            // ═══════════════════════════════════════════════════════════

            _lblZoomInfo = new Label();
            _lblZoomInfo.Text = "Zoom: 100%  |  Rotación: 0°";
            _lblZoomInfo.ForeColor = Color.White;
            _lblZoomInfo.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            _lblZoomInfo.AutoSize = true;
            _lblZoomInfo.Location = new Point(1250, 17);   // ← Movido a la derecha del botón Cerrar

            panelControles.Controls.Add(btnRotarIzq);
            panelControles.Controls.Add(btnRotarDer);
            panelControles.Controls.Add(btnZoomIn);
            panelControles.Controls.Add(btnZoomOut);
            panelControles.Controls.Add(btnReset);
            panelControles.Controls.Add(btnCerrar);       // ← Agregar botón cerrar
            panelControles.Controls.Add(_lblZoomInfo);

            // ═══ Panel inferior con instrucciones ═══
            Label lblInstrucciones = new Label();
            lblInstrucciones.Text = "🖱️ Rueda del mouse: Zoom  |  ⌨️ Flechas ← →: Rotar  |  ESC: Cerrar";
            lblInstrucciones.Dock = DockStyle.Bottom;
            lblInstrucciones.Height = 30;
            lblInstrucciones.BackColor = Color.FromArgb(40, 40, 40);
            lblInstrucciones.ForeColor = Color.LightGray;
            lblInstrucciones.Font = new Font("Segoe UI", 9, FontStyle.Italic);
            lblInstrucciones.TextAlign = ContentAlignment.MiddleCenter;

            // ═══ Eventos de teclado ═══
            _formImagenAmpliada.KeyDown += (s, ev) =>
            {
                switch (ev.KeyCode)
                {
                    case Keys.Escape:
                    case Keys.Space:
                        CerrarImagenAmpliada();
                        break;
                    case Keys.Left:
                        _rotacionAmpliado -= 90f;
                        AplicarTransformacionAmpliada();
                        break;
                    case Keys.Right:
                        _rotacionAmpliado += 90f;
                        AplicarTransformacionAmpliada();
                        break;
                    case Keys.Up:
                    case Keys.Add:
                    case Keys.Oemplus:
                        _zoomAmpliado = Math.Min(_zoomAmpliado + 0.25f, 5.0f);
                        AplicarTransformacionAmpliada();
                        break;
                    case Keys.Down:
                    case Keys.Subtract:
                    case Keys.OemMinus:
                        _zoomAmpliado = Math.Max(_zoomAmpliado - 0.25f, 0.25f);
                        AplicarTransformacionAmpliada();
                        break;
                }
            };

            // ═══ Rueda del mouse para zoom ═══
            MouseEventHandler zoomHandler = (s, ev) =>
            {
                if (ev.Delta > 0)
                    _zoomAmpliado = Math.Min(_zoomAmpliado + 0.15f, 5.0f);
                else
                    _zoomAmpliado = Math.Max(_zoomAmpliado - 0.15f, 0.25f);

                AplicarTransformacionAmpliada();
            };

            _pbAmpliada.MouseWheel += (s, ev) => zoomHandler(s, ev);
            _panelScrollAmpliada.MouseWheel += (s, ev) => zoomHandler(s, ev);

            // ═══ Agregar controles ═══
            _formImagenAmpliada.Controls.Add(_panelScrollAmpliada);
            _formImagenAmpliada.Controls.Add(panelControles);
            _formImagenAmpliada.Controls.Add(lblInstrucciones);

            // ═══ Tamaño y posición ═══
            Rectangle screenBounds = Screen.PrimaryScreen.WorkingArea;
            int width = (int)(screenBounds.Width * 0.95);
            int height = (int)(screenBounds.Height * 0.95);
            _formImagenAmpliada.Size = new Size(width, height);
            _formImagenAmpliada.Location = new Point(
                (screenBounds.Width - width) / 2,
                (screenBounds.Height - height) / 2
            );

            // ═══ FormClosing para limpiar recursos ═══
            _formImagenAmpliada.FormClosing += (s, ev) => LimpiarRecursosAmpliada();

            _imagenAmpliadaVisible = true;
            _formImagenAmpliada.Show();

            AplicarTransformacionAmpliada();
            _formImagenAmpliada.Focus();
        }
        /// <summary>
        /// Crea un botón estilizado para el panel de controles
        /// </summary>
        private Button CrearBotonControl(string texto, int x, int y)
        {
            Button btn = new Button();
            btn.Text = texto;
            btn.Location = new Point(x, y);
            btn.Size = new Size(180, 34);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            btn.BackColor = Color.FromArgb(60, 60, 60);
            btn.ForeColor = Color.White;
            btn.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            return btn;
        }
        /// <summary>
        /// 🔥 Aplica la rotación y el zoom a la imagen ampliada
        /// </summary>
        private void AplicarTransformacionAmpliada()
        {
            if (_imagenOriginalAmpliada == null || _pbAmpliada == null) return;

            try
            {
                int anchoBase = _imagenOriginalAmpliada.Width;
                int altoBase = _imagenOriginalAmpliada.Height;

                // ─── Normalizar rotación al rango [0, 360) ───
                float rot = _rotacionAmpliado % 360f;
                if (rot < 0) rot += 360f;

                // ─── Calcular dimensiones después de rotación ───
                double rad = rot * Math.PI / 180.0;
                double cos = Math.Abs(Math.Cos(rad));
                double sin = Math.Abs(Math.Sin(rad));

                int anchoRotado = (int)Math.Round(anchoBase * cos + altoBase * sin);
                int altoRotado = (int)Math.Round(anchoBase * sin + altoBase * cos);

                // ─── Aplicar zoom ───
                int anchoFinal = (int)Math.Round(anchoRotado * _zoomAmpliado);
                int altoFinal = (int)Math.Round(altoRotado * _zoomAmpliado);

                // ─── Limitar tamaño máximo (evitar OutOfMemory) ───
                const int MAX_DIM = 10000;
                if (anchoFinal > MAX_DIM || altoFinal > MAX_DIM)
                {
                    float factor = Math.Min((float)MAX_DIM / anchoFinal, (float)MAX_DIM / altoFinal);
                    anchoFinal = (int)(anchoFinal * factor);
                    altoFinal = (int)(altoFinal * factor);
                }

                if (anchoFinal <= 0 || altoFinal <= 0) return;

                // ─── Crear bitmap transformado ───
                Bitmap bmpNuevo = new Bitmap(anchoFinal, altoFinal, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(bmpNuevo))
                {
                    g.Clear(Color.Black);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    // Centro del bitmap destino
                    float centroDestinoX = anchoFinal / 2f;
                    float centroDestinoY = altoFinal / 2f;

                    // Centro de la imagen original
                    float centroOrigenX = anchoBase / 2f;
                    float centroOrigenY = altoBase / 2f;

                    // Transformaciones: mover al centro, rotar, escalar, mover de vuelta
                    g.TranslateTransform(centroDestinoX, centroDestinoY);
                    g.RotateTransform(rot);
                    g.ScaleTransform(_zoomAmpliado, _zoomAmpliado);
                    g.TranslateTransform(-centroOrigenX, -centroOrigenY);

                    // Dibujar la imagen original completa
                    g.DrawImage(_imagenOriginalAmpliada, 0, 0, anchoBase, altoBase);
                }

                // ─── Guardar referencia anterior para liberar después ───
                Image imagenAnterior = _pbAmpliada.Image;

                // ─── Asignar nueva imagen ───
                _pbAmpliada.Image = bmpNuevo;
                _pbAmpliada.Size = new Size(anchoFinal, altoFinal);

                // 🔥 CLAVE: Liberar la imagen anterior DESPUÉS de asignar la nueva,
                // y solo si NO es la original (evitar disponer _imagenOriginalAmpliada)
                if (imagenAnterior != null && imagenAnterior != _imagenOriginalAmpliada)
                {
                    try { imagenAnterior.Dispose(); } catch { /* ignorar */ }
                }

                // ─── Actualizar etiqueta ───
                if (_lblZoomInfo != null)
                {
                    _lblZoomInfo.Text = $"Zoom: {(_zoomAmpliado * 100):F0}%  |  Rotación: {rot:F0}°";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en transformación: {ex.Message}");
                MessageBox.Show($"Error al transformar imagen: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// 🔥 Libera los recursos de la ventana ampliada. Se llama al cerrar el form.
        /// </summary>
        private void LimpiarRecursosAmpliada()
        {
            try
            {
                // Primero, desasociar la imagen del PictureBox para que no intente dibujarla
                if (_pbAmpliada != null)
                {
                    Image img = _pbAmpliada.Image;
                    _pbAmpliada.Image = null;

                    // Liberar el bitmap generado (NO la original)
                    if (img != null && img != _imagenOriginalAmpliada)
                    {
                        try { img.Dispose(); } catch { }
                    }
                }

                // Liberar la imagen original
                if (_imagenOriginalAmpliada != null)
                {
                    try { _imagenOriginalAmpliada.Dispose(); } catch { }
                    _imagenOriginalAmpliada = null;
                }

                _pbAmpliada = null;
                _panelScrollAmpliada = null;
                _lblZoomInfo = null;

                // Resetear estado
                _zoomAmpliado = 1.0f;
                _rotacionAmpliado = 0f;
            }
            catch { /* ignorar errores al limpiar */ }
        }

        private void CerrarImagenAmpliada()
        {
            if (_formImagenAmpliada != null && !_formImagenAmpliada.IsDisposed)
            {
                _formImagenAmpliada.Close();   // Dispara FormClosing → LimpiarRecursosAmpliada
                _formImagenAmpliada = null;
            }
            _imagenAmpliadaVisible = false;
        }

        private void listBoxMiniaturas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxMiniaturas.SelectedIndex >= 0 && listBoxMiniaturas.SelectedIndex < _rutasImagenes.Count)
            {
                _imagenActualIndex = listBoxMiniaturas.SelectedIndex;
                MostrarImagenActual();
                btnAnterior.Enabled = _imagenActualIndex > 0;
                btnSiguiente.Enabled = _imagenActualIndex < _rutasImagenes.Count - 1;
            }
        }

        private void FormVerImagenes_FormClosing(object sender, FormClosingEventArgs e)
        {
            CerrarImagenAmpliada();
        }
    }
}