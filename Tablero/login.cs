using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tablero
{
    public partial class login : Form
    {
        // Variables para la animación de imágenes
        private Image image1;
        private Image image2;
        private Image image3;
        private Timer animationTimer;
        private float transitionProgress = 0f;
        private const float TransitionSpeed = 0.1f;
        private bool isHovering = false;
        private bool showingImage3 = false;
        //variable para la conexión a la base de datos
        // string connectionString = "Host=localhost;Username=postgres;Password=Picolargo789;Database=Reporteo";
        string connectionString = string.Empty;


        // seccion para activar el drag en el formulario
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

        private bool mostrarPassword = false;

        public login()
        {
            InitializeComponent();
            InitializeAnimation();// Inicializa el temporizador de animación

            // Configuración para lbl_limpiar
            LabelAnimator.SetupLabel(
                lbl_limpiar,
                normalColor: Color.FromArgb(25, 118, 210),
                hoverColor: Color.FromArgb(129, 212, 250),
                clickColor: Color.FromArgb(33, 150, 243),
                originalSize: 9.75f,
                clickedSize: 8.75f);
            // Configuración para lbl_salir (mismos parámetros)
            LabelAnimator.SetupLabel(
                lbl_salir,
                normalColor: Color.FromArgb(25, 118, 210),
                hoverColor: Color.FromArgb(129, 212, 250),
                clickColor: Color.FromArgb(33, 150, 243),
                originalSize: 9.75f,
                clickedSize: 8.75f);

            // Asegúrate de cargar tus imágenes desde los recursos
            image1 = Properties.Resources._5172968_disable_eye_hidden_hide_internet_icon; // Reemplaza con tus imágenes reales
            image2 = Properties.Resources._5173015_eye_focus_internet_scan_security_icon;
            image3 = Properties.Resources._5172950_business_eye_focus_internet_security_icon;

            this.FormBorderStyle = FormBorderStyle.None;
            this.MouseDown += (sender, e) => {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(this.Handle, 0x112, 0xf012, 0);
                }
            };
        }

        private void InitializeAnimation()
        {
            animationTimer = new Timer
            {
                Interval = 16 // ~60 FPS
            };
            animationTimer.Tick += AnimationTimer_Tick;
            animationTimer.Start();
        }

        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            if (isHovering)
            {
                transitionProgress += TransitionSpeed;
                if (transitionProgress > 1f) transitionProgress = 1f;
            }
            else
            {
                transitionProgress -= TransitionSpeed;
                if (transitionProgress < 0f) transitionProgress = 0f;
            }

            if (showingImage3)
            {
                // Animación entre imagen2 e imagen3
                if (transitionProgress > 0f && transitionProgress < 1f)
                {
                    pictureBox4.Image = BlendImages(image3, image2, transitionProgress);
                }
                else if (transitionProgress <= 0f)
                {
                    pictureBox4.Image = image3;
                }
                else if (transitionProgress >= 1f)
                {
                    pictureBox4.Image = image2;
                }
            }
            else
            {
                // Animación entre imagen1 e imagen2
                if (transitionProgress > 0f && transitionProgress < 1f)
                {
                    pictureBox4.Image = BlendImages(image1, image2, transitionProgress);
                }
                else if (transitionProgress <= 0f)
                {
                    pictureBox4.Image = image1;
                }
                else if (transitionProgress >= 1f)
                {
                    pictureBox4.Image = image2;
                }
            }
        }

        private Bitmap BlendImages(Image img1, Image img2, float blendFactor)
        {
            if (img1 == null || img2 == null) return null;

            Bitmap bmp = new Bitmap(img1.Width, img1.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Primero dibujar la imagen1 con opacidad decreciente
                ColorMatrix cm1 = new ColorMatrix();
                cm1.Matrix33 = 1 - blendFactor; // Opacidad
                ImageAttributes ia1 = new ImageAttributes();
                ia1.SetColorMatrix(cm1);

                g.DrawImage(img1,
                    new Rectangle(0, 0, img1.Width, img1.Height),
                    0, 0, img1.Width, img1.Height,
                    GraphicsUnit.Pixel, ia1);

                // Luego dibujar la imagen2 con opacidad creciente
                ColorMatrix cm2 = new ColorMatrix();
                cm2.Matrix33 = blendFactor; // Opacidad
                ImageAttributes ia2 = new ImageAttributes();
                ia2.SetColorMatrix(cm2);

                g.DrawImage(img2,
                    new Rectangle(0, 0, img2.Width, img2.Height),
                    0, 0, img2.Width, img2.Height,
                    GraphicsUnit.Pixel, ia2);
            }
            return bmp;
        }

        private void pictureBox4_MouseEnter(object sender, EventArgs e)
        {
            isHovering = true;
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            isHovering = false;
        }

        private void btn_iniciar_Click(object sender, EventArgs e)
        {
            //Tablero.Properties.Settings.Default.Servidor = string.Empty;
            //Tablero.Properties.Settings.Default.Usuario = string.Empty;
            //Tablero.Properties.Settings.Default.Contrasena = string.Empty;
            //Tablero.Properties.Settings.Default.Save();
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
            }
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            bool isValid = dbHelper.ValidateUser(txt_user_name.Text, txt_password.Text);

            if (isValid)
            {
                // Obtener información adicional del usuario
                DataRow userInfo = dbHelper.GetUserInfo(txt_user_name.Text);

                if (userInfo != null)
                {
                    int idUser = Convert.ToInt32(userInfo["ID_User"]);
                    string nivel = userInfo["Nivel"].ToString();
                    string noEmpleado = userInfo["No_Empleado"].ToString();

                    // Usuario válido
                    this.Visible = false;
                    Form_principal principal = new Form_principal(noEmpleado, txt_user_name.Text, idUser, nivel, connectionString);
                    principal.WindowState = FormWindowState.Maximized; // <-- Aquí fuerzas el modo maximizado
                    principal.Show(); // Muestra el formulario principal
                    // Guardar en sesión o usar esta información
                }
            }
            else
            {
                // Credenciales incorrectas
                MetroFramework.MetroMessageBox.Show(this, "El usuario y/o la contraseña son incorrectos. Por favor, verifique sus datos e intente nuevamente.\r\n\r\n", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                limpiarCampos(); // Limpia los campos de texto
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            // Alternar entre image1 e image3 con cada clic
            showingImage3 = !showingImage3;

            // Reiniciar la animación
            transitionProgress = 0f;

            // Forzar una actualización inmediata
            if (showingImage3 ||!(mostrarPassword))
            {
                pictureBox4.Image = image3;
                // Mostrar contraseña
                txt_password.PasswordChar = '\0'; // Esto quita el carácter de ocultamiento
                mostrarPassword = true;
            }
            else
            {
                pictureBox4.Image = image1;
                // Ocultar contraseña
                txt_password.PasswordChar = '*';
                mostrarPassword = false;
            }
        }

        private void lbl_limpiar_Click(object sender, EventArgs e)
        {
            limpiarCampos(); // Llama al método para limpiar los campos de texto
        }
        private void limpiarCampos()
        {
            txt_user_name.Clear();
            txt_password.Clear();
            txt_user_name.Focus(); // Enfoca el campo de usuario
        }

        private void lbl_salir_Click(object sender, EventArgs e)
        {
            Application.Exit(); // Cierra la aplicación
        }

        private void txt_password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_iniciar_Click(btn_iniciar, EventArgs.Empty);
                e.SuppressKeyPress = true; // evita el beep
            }
        }

        private void login_Shown(object sender, EventArgs e)
        {

            //Tablero.Properties.Settings.Default.Servidor = string.Empty;
            //Tablero.Properties.Settings.Default.Usuario = string.Empty;
            //Tablero.Properties.Settings.Default.Contrasena = string.Empty;
            //Tablero.Properties.Settings.Default.Save();
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
            }
        }
    }
}
