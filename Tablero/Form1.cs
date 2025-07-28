using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tablero
{
    public partial class Form_principal : MaterialForm
    {
        private string var_no_empledo = string.Empty;
        private string var_nom_empledo = string.Empty;
        public Form_principal(string var_no_empledo, string var_nom_empledo)
        {
            InitializeComponent();
            this.var_no_empledo = var_no_empledo;// Asignar el número de empleado al campo correspondiente
            this.var_nom_empledo = var_nom_empledo;// Asignar el nombre del empleado al campo correspondiente

            lbl_no_emp2.Text = var_no_empledo; // Mostrar el número de empleado en el label correspondiente
            lbl_nom2.Text = var_nom_empledo; // Mostrar el nombre del empleado en el label correspondiente

            // Initialize MaterialSkinManager and set the theme and color scheme  
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Orange600, Primary.Orange600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);

            // Forzar el color después de inicializar MaterialSkin  
            pnl_principal.BackColor = Color.FromArgb(192, 255, 192);
            pnl_Metahora.BackColor = Color.FromArgb(192, 255, 192);
            pnl_ftt.BackColor = Color.FromArgb(255, 255, 128);

            // Personalización de dgv_mecanico
            dgv_mecanico.EnableHeadersVisualStyles = false;
            dgv_mecanico.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_mecanico.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_mecanico.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_mecanico.BackgroundColor = Color.White; // Fondo blanco
            dgv_mecanico.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_mecanico.DefaultCellStyle.ForeColor = Color.Black;
            dgv_mecanico.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_mecanico.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_mecanico.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_mecanico.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_mecanico.RowHeadersVisible = false;
            dgv_mecanico.AllowUserToAddRows = true;
            dgv_mecanico.AllowUserToDeleteRows = true;
            dgv_mecanico.AllowUserToResizeRows = false;
            dgv_mecanico.ReadOnly = false;
            dgv_mecanico.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_mecanico.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_mecanico.MultiSelect = false;
            dgv_mecanico.BorderStyle = BorderStyle.None;

            // Personalización de dgv_almacen
            dgv_almacen.EnableHeadersVisualStyles = false;
            dgv_almacen.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_almacen.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_almacen.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_almacen.BackgroundColor = Color.White; // Fondo blanco
            dgv_almacen.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_almacen.DefaultCellStyle.ForeColor = Color.Black;
            dgv_almacen.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_almacen.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_almacen.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_almacen.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_almacen.RowHeadersVisible = false;
            dgv_almacen.AllowUserToAddRows = true;
            dgv_almacen.AllowUserToDeleteRows = true;
            dgv_almacen.AllowUserToResizeRows = false;
            dgv_almacen.ReadOnly = false;
            dgv_almacen.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_almacen.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_almacen.MultiSelect = false;
            dgv_almacen.BorderStyle = BorderStyle.None;

            // Personalización de dgv_operativo
            dgv_operativo.EnableHeadersVisualStyles = false;
            dgv_operativo.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_operativo.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_operativo.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_operativo.BackgroundColor = Color.White; // Fondo blanco
            dgv_operativo.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_operativo.DefaultCellStyle.ForeColor = Color.Black;
            dgv_operativo.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_operativo.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_operativo.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_operativo.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_operativo.RowHeadersVisible = false;
            dgv_operativo.AllowUserToAddRows = true;
            dgv_operativo.AllowUserToDeleteRows = true;
            dgv_operativo.AllowUserToResizeRows = false;
            dgv_operativo.ReadOnly = false;
            dgv_operativo.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_operativo.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_operativo.MultiSelect = false;
            dgv_operativo.BorderStyle = BorderStyle.None;

            // Personalización de dgv_calidad
            dgv_calidad.EnableHeadersVisualStyles = false;
            dgv_calidad.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_calidad.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_calidad.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_calidad.BackgroundColor = Color.White; // Fondo blanco
            dgv_calidad.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_calidad.DefaultCellStyle.ForeColor = Color.Black;
            dgv_calidad.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_calidad.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_calidad.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_calidad.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_calidad.RowHeadersVisible = false;
            dgv_calidad.AllowUserToAddRows = true;
            dgv_calidad.AllowUserToDeleteRows = true;
            dgv_calidad.AllowUserToResizeRows = false;
            dgv_calidad.ReadOnly = false;
            dgv_calidad.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dgv_calidad.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv_calidad.MultiSelect = false;
            dgv_calidad.BorderStyle = BorderStyle.None;

            // Personalización de dgv_users
            dgv_users.EnableHeadersVisualStyles = false;
            dgv_users.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_users.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_users.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_users.BackgroundColor = Color.White; // Fondo blanco
            dgv_users.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_users.DefaultCellStyle.ForeColor = Color.Black;
            dgv_users.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_users.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_users.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_users.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_users.RowHeadersVisible = false;
            dgv_users.BorderStyle = BorderStyle.None;

            // Personalización de dgv_metas
            dgv_metas.EnableHeadersVisualStyles = false;
            dgv_metas.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv_metas.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dgv_metas.BackgroundColor = Color.White; // Fondo blanco
            dgv_metas.DefaultCellStyle.BackColor = Color.White; // Renglones blancos
            dgv_metas.DefaultCellStyle.ForeColor = Color.Black;
            dgv_metas.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv_metas.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv_metas.DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            dgv_metas.GridColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv_metas.RowHeadersVisible = false;
            dgv_metas.BorderStyle = BorderStyle.None;
            
        }

        private void dgv_mecanico_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_mecanico.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void SoloNumeros_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite solo dígitos y teclas de control (como backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void dgv_almacen_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_almacen.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void dgv_operativo_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_operativo.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void dgv_calidad_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_calidad.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Quita la restricción en otras columnas
                }
            }
        }

        private void txt_Tiempo_comida_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite solo dígitos y teclas de control (como backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
    }
}
