using System;
using System.Windows.Forms;

namespace Tablero
{
    partial class Automatico
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
            this.SuspendLayout();
            // 
            // Automatico
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(200, 100);
            this.Name = "Automatico";
            this.Text = "Servicio Automático Tablero";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.ResumeLayout(false);
        }
        private void AbrirConfiguracionEmail()
        {
            try
            {

                string var_servidor = Tablero.Properties.Settings.Default.Servidor;
                string var_password = Tablero.Properties.Settings.Default.Contrasena;
                string var_usuario = Tablero.Properties.Settings.Default.Usuario;

                if (!string.IsNullOrEmpty(var_servidor) && !string.IsNullOrEmpty(var_password) && !string.IsNullOrEmpty(var_usuario))
                {
                    // Construir la cadena de conexión
                    connectionString = $"Host={var_servidor};Port=5433;Username={var_usuario};Password={var_password};Database=Reporteo";
                }
                else
                {
                    // Si no hay configuración guardada, abrir el formulario de conexión
                    var frm = new Coneccion_Server();
                    frm.Owner = this;      // <-- importante
                    frm.TopMost = true;
                    frm.Show();
                    this.Hide();
                    return; // Salir aquí para no continuar
                }
                // Verificar si el formulario ya está abierto
                Form existingForm = Application.OpenForms["Form_principal"];
                if (existingForm != null)
                {
                    existingForm.WindowState = FormWindowState.Normal;
                    existingForm.BringToFront();
                    return;
                }

                // Crear y mostrar el formulario principal (sin parámetros)
                Form_principal principal = new Form_principal("1", "Admin", 1, "Super Administrador", connectionString);
                principal.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir configuración de email: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}