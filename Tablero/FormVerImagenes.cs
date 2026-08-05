using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

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

            _formImagenAmpliada = new Form();
            _formImagenAmpliada.FormBorderStyle = FormBorderStyle.None;
            _formImagenAmpliada.BackColor = Color.Black;
            _formImagenAmpliada.StartPosition = FormStartPosition.CenterScreen;
            _formImagenAmpliada.TopMost = true;
            _formImagenAmpliada.KeyPreview = true;

            PictureBox pbAmpliada = new PictureBox();
            pbAmpliada.Dock = DockStyle.Fill;
            pbAmpliada.Image = (Image)pictureBoxImagen.Image.Clone();
            pbAmpliada.SizeMode = PictureBoxSizeMode.Zoom;
            pbAmpliada.BackColor = Color.Black;

            _formImagenAmpliada.Controls.Add(pbAmpliada);

            Rectangle screenBounds = Screen.PrimaryScreen.WorkingArea;
            int width = (int)(screenBounds.Width * 0.9);
            int height = (int)(screenBounds.Height * 0.9);
            _formImagenAmpliada.Size = new Size(width, height);
            _formImagenAmpliada.Location = new Point(
                (screenBounds.Width - width) / 2,
                (screenBounds.Height - height) / 2
            );

            _formImagenAmpliada.Click += (s, ev) => CerrarImagenAmpliada();
            _formImagenAmpliada.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Escape || ev.KeyCode == Keys.Space)
                    CerrarImagenAmpliada();
            };
            pbAmpliada.Click += (s, ev) => CerrarImagenAmpliada();

            _imagenAmpliadaVisible = true;
            _formImagenAmpliada.Show();
            _formImagenAmpliada.Focus();
        }

        private void CerrarImagenAmpliada()
        {
            if (_formImagenAmpliada != null && !_formImagenAmpliada.IsDisposed)
            {
                _formImagenAmpliada.Close();
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