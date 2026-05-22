using ClosedXML.Excel;
using LiveCharts;
using MaterialSkin;
using MaterialSkin.Controls;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telerik.WinControls.UI;
//using System.Windows.Forms.DataVisualization.Charting;


namespace Tablero
{
    public partial class Form_principal : MaterialForm
    {
        private string id_global_users = string.Empty; // Variable para almacenar el ID del usuario seleccionado en el DataGridView
        private string id_global_meta_deshidratado = string.Empty;
        private string id_global_costo = string.Empty;
        private string id_global_meta_empacado = string.Empty;
        private string id_global_meta_inspec = string.Empty;
        private string id_global_meta_evaporado = string.Empty;
        private string id_global_meta_grind = string.Empty;
        private string id_global_meta_revolturas = string.Empty;
        private string id_global_meta_polvos = string.Empty;
        private string id_global_meta_maquinas = string.Empty;
        private string id_global_ficha = string.Empty;
        private string id_global_polvos_calidad = string.Empty;
        private string id_global_tunel_calidad = string.Empty;
        private string id_global_detalles_OP = string.Empty;
        private int id_user = 0;
        private string nivel_user = string.Empty;
        private string nombre_admin = string.Empty;
        private string numero_empleado_admin = string.Empty;
        private int id_supervisor_global = 0;
        private bool filtroUsuariosActivo = false;
        private bool filtroUsuariosActivo_OP = false;
        private bool filtrodetallesOP = false;
        private bool filtroreporte_CD = false;
        private bool filtroreporte_CO = false;
        private int filtroUsuariosActivo_OP_dgv_activo = 0;
        private string horaInicioText = string.Empty;
        private string horaFinText = string.Empty;
        private DateTime horaInicio;
        private DateTime horaFin;
        private bool editar = false;
        private bool borrar = false;
        //variable para la conexión a la base de datos
        string connectionString = string.Empty;
        // Variables para la animación de las imágenes
        private Image image1;
        private Image image2;
        private Image image3;
        private bool showingImage3 = false;
        private float transitionProgress = 0f;
        private const float TransitionSpeed = 0.1f;
        private bool mostrarPassword = false;
        private bool isHovering = false;
        private Timer animationTimer;
        // Fin de variables para la animación de las imágenes
        // Variables para el envío de correos
        private string servidor_smtp = string.Empty;
        private string RemitenteEMail = string.Empty;
        private string PasswordEmail = string.Empty;
        private int PuertoSMTP = 0;
        private string DestinatariosEmail = string.Empty;
        private bool SSLCheck = false;
        // Fin de variables para el envío de correos
        List<ValueTuple<string, string, string>> valores_mecanico =
        new List<(string, string, string)>();
        List<ValueTuple<string, string, string>> valores_operativos =
        new List<(string, string, string)>();
        private Editar _editarForm = null;
        private string Kg_pz_str = string.Empty;
        public Form_principal(string var_no_empledo, string var_nom_empledo, int ID_usuario, string nivel, string conexionstring)
        {
            InitializeComponent();
            // Configurar propiedades similares al ejemplo
            rgv_reporte_consolidado.EnableGrouping = false;
            rgv_reporte_consolidado.EnableHotTracking = true;
            rgv_reporte_consolidado.ShowFilteringRow = false;
            rgv_reporte_consolidado.EnableFiltering = true;
            rgv_reporte_consolidado.EnableCustomFiltering = true;

            // Configurar propiedades similares al ejemplo
           // rgv_reporte_costo.EnableGrouping = false;
            rgv_reporte_costo.EnableHotTracking = true;
            rgv_reporte_costo.ShowFilteringRow = false;
            rgv_reporte_costo.EnableFiltering = true;
            rgv_reporte_costo.EnableCustomFiltering = true;

            // Configurar propiedades similares al ejemplo
            rgv_reporte_Tiempos.EnableGrouping = false;
            rgv_reporte_Tiempos.EnableHotTracking = true;
            rgv_reporte_Tiempos.ShowFilteringRow = false;
            rgv_reporte_Tiempos.EnableFiltering = true;
            rgv_reporte_Tiempos.EnableCustomFiltering = true;


            this.WindowState = FormWindowState.Maximized;

            lbl_no_emp2.Text = var_no_empledo; // Mostrar el número de empleado en el label correspondiente
            lbl_nom2.Text = var_nom_empledo.ToUpper(); // Mostrar el nombre del empleado en el label correspondiente
            connectionString = conexionstring; // Asignar la cadena de conexión pasada como parámetro
            id_user = ID_usuario; // Asignar el ID del usuario pasado como parámetro
            nivel_user = nivel; // Asignar el nivel del usuario pasado como parámetro

            nombre_admin = var_nom_empledo.ToUpper(); // Almacenar el nombre del usuario para uso futuro
            numero_empleado_admin = var_no_empledo; // Almacenar el número de empleado para uso futuro

            // Initialize MaterialSkinManager and set the theme and color scheme  
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Orange600, Primary.Orange600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);


            // Personaliza todos los DataGridView de una sola línea cada uno
            PersonalizarDataGridView(dgv_mecanico, true);
            PersonalizarDataGridView(dgv_operativo, true);
            PersonalizarDataGridView(dgv_users);
            PersonalizarDataGridView(dgv_metas_des);
            PersonalizarDataGridView(dgv_metas_emp);
            PersonalizarDataGridView(dgv_metas_insp);
            PersonalizarDataGridView(dgv_metas_Eva);
            PersonalizarDataGridView(dgv_metas_Grind);
            PersonalizarDataGridView(dgv_metas_revolturas);
            PersonalizarDataGridView(dgv_metas_polvos);
            PersonalizarDataGridView(dgv_metas_maquinas);
            PersonalizarDataGridView(dgv_polvos_calidad);
            PersonalizarDataGridView(dgv_detalles_op);
            PersonalizarDataGridView(dgv_Tunel_calidad);
            PersonalizarDataGridView(dgv_reporte_merma);
            PersonalizarDataGridView(dgv_reporte_concentrado);
            PersonalizarDataGridView(dgv_reporte_concentrado_otras);
            PersonalizarDataGridView(dgv_reporte_merma_S);
            InitializeAnimation();// Inicializa el temporizador de animación
            // Asegúrate de cargar tus imágenes desde los recursos
            image1 = Properties.Resources._5172968_disable_eye_hidden_hide_internet_icon;
            image2 = Properties.Resources._5173015_eye_focus_internet_scan_security_icon;
            image3 = Properties.Resources._5172950_business_eye_focus_internet_security_icon;
        }
        //Animacion de boton mostrar contraseña
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

        private void PersonalizarDataGridView(DataGridView dgv, bool editable = false)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 152, 0); // Naranja
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold);

            dgv.BackgroundColor = Color.White;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243); // Azul Material
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Regular);

            dgv.GridColor = Color.FromArgb(255, 152, 0);
            dgv.RowHeadersVisible = false;
            dgv.BorderStyle = BorderStyle.None;

            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;

            // Configuración para permitir redimensionamiento
            dgv.AllowUserToResizeColumns = true;
            dgv.AllowUserToResizeRows = false; // Normalmente no es necesario redimensionar filas
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing; // ¡Importante!

            // Opcional según si quieres permitir edición
            dgv.AllowUserToAddRows = editable;
            dgv.AllowUserToDeleteRows = editable;
            dgv.ReadOnly = !editable;
        }

        private void dgv_mecanico_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgv_mecanico.CurrentCell.ColumnIndex == 0) // Primera columna
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress; // Evita múltiples suscripciones
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
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
        private void dgv_operativo_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // 1. Validación para solo números en columna 0
            if (dgv_operativo.CurrentCell.ColumnIndex == 0)
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress;
                    tb.KeyPress += SoloNumeros_KeyPress;
                }
            }
            else
            {
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
                if (tb != null)
                {
                    tb.KeyPress -= SoloNumeros_KeyPress;
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
        // Configuración básica del ListView
        private void ConfigurarListViewSimple()
        {
            // Configurar vista como lista (solo texto, sin columnas)
            listView1.View = View.List;

            // Permitir selección múltiple
            listView1.MultiSelect = true;

            // Habilitar checkboxes para selección más intuitiva
            listView1.CheckBoxes = true;

            // Opcional: Permitir selección completa del item
            listView1.FullRowSelect = true;

            // Opcional: Eliminar líneas de cuadrícula (más limpio)
            listView1.GridLines = false;
        }

        // Cargar las semanas desde la 1 hasta la semana actual
        private void CargarSemanasSimple()
        {
            // Limpiar items existentes
            listView1.Items.Clear();

            // Obtener fecha actual
            DateTime fechaActual = DateTime.Now;

            // Obtener número de semana actual (semana empieza en lunes)
            int semanaActual = ObtenerNumeroSemana(fechaActual);

            // Cargar semanas del 1 al número de semana actual
            for (int semana = 1; semana <= semanaActual; semana++)
            {
                // Crear el item con el número de semana
                ListViewItem item = new ListViewItem(semana.ToString());

                // Guardar el número de semana en Tag para usarlo después
                item.Tag = semana;

                // Agregar al ListView
                listView1.Items.Add(item);
            }

            // Opcional: Seleccionar todas las semanas por defecto
            foreach (ListViewItem item in listView1.Items)
            {
                item.Selected = false;  // Selecciona visualmente
                item.Checked = false;    // Marca el checkbox
            }
        }
        // Cargar las semanas desde la 1 hasta la semana actual
        private void CargartodasSemanas()
        {
            // Limpiar items existentes
            listView1.Items.Clear();

            // Cargar semanas del 1 al número de semana actual
            for (int semana = 1; semana <= 52; semana++)
            {
                // Crear el item con el número de semana
                ListViewItem item = new ListViewItem(semana.ToString());

                // Guardar el número de semana en Tag para usarlo después
                item.Tag = semana;

                // Agregar al ListView
                listView1.Items.Add(item);
            }

            // Opcional: Seleccionar todas las semanas por defecto
            foreach (ListViewItem item in listView1.Items)
            {
                item.Selected = false;  // Selecciona visualmente
                item.Checked = false;    // Marca el checkbox
            }
        }

        // Método para obtener el número de semana (Lunes a Domingo)
        private int ObtenerNumeroSemana(DateTime fecha)
        {
            // Usar cultura invariable
            CultureInfo culture = CultureInfo.InvariantCulture;
            Calendar calendar = culture.Calendar;

            // Semana comienza en LUNES, usando estándar ISO
            return calendar.GetWeekOfYear(fecha, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        // Método para obtener las semanas seleccionadas (usando selección normal)
        private List<int> ObtenerSemanasSeleccionadas()
        {
            List<int> semanas = new List<int>();

            foreach (ListViewItem item in listView1.SelectedItems)
            {
                // Obtener el número de semana del Tag o del texto
                if (item.Tag != null)
                {
                    semanas.Add((int)item.Tag);
                }
                else
                {
                    semanas.Add(int.Parse(item.Text));
                }
            }

            return semanas;
        }

        // Método para obtener las semanas seleccionadas (usando checkboxes)
        private List<int> ObtenerSemanasSeleccionadasConCheckbox()
        {
            List<int> semanas = new List<int>();

            foreach (ListViewItem item in listView1.Items)
            {
                if (item.Checked)
                {
                    if (item.Tag != null)
                    {
                        semanas.Add((int)item.Tag);
                    }
                    else
                    {
                        semanas.Add(int.Parse(item.Text));
                    }
                }
            }

            return semanas;
        }
        private void Form_principal_Load(object sender, EventArgs e)
        {
            materialTabControl1.TabPages.Remove(tabPage36);// eliminar la pestaña de productos hasta programar ese modulo es el modulo de productos o fichas de producción
            
            if (nivel_user == "Super Administrador")
            {
                lbl_user_no_emp.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                lbl_Nom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                dtp_polvos.Value = DateTime.Now;
                dtp_tunel.Value = DateTime.Now;
                string var_servidor = Tablero.Properties.Settings.Default.Servidor;
                string nombreEquipo = Environment.MachineName;
                if (nombreEquipo.ToUpper() == var_servidor.ToUpper())
                {
                    materialTabControl1.TabPages.Remove(tabPage1);
                    materialTabControl1.TabPages.Remove(tabPage2);
                    materialTabControl1.TabPages.Remove(tabPage10);
                    materialTabControl1.TabPages.Remove(tabPage4);
                    materialTabControl1.TabPages.Remove(tabPage9);
                    materialTabControl1.TabPages.Remove(tabPage11);
                    materialCardtab_users.Visible = false;
                    materialCard_users.Visible = false;
                    materialCard18.Location = new Point(14, 14);
                    emaildatos();
                    email_varibles();
                }
                else 
                {
                    actualiza_grid_users(); // Llamar al método para actualizar el DataGridView de usuarios
                    emaildatos();
                    email_varibles();

                    // Obtener y mostrar el número de semana inicial
                    ActualizarNumeroSemana();
                    actualiza_grid_Deshitratado();
                    actualiza_grid_Empacado();
                    actualiza_grid_inspec();
                    actualiza_grid_evaporado();
                    actualiza_grid_grind();
                    actualiza_revolturas();
                    actualiza_maquinas();
                    actualiza_polvos();
                    actualiza_polvos_calidad();
                    actualiza_detalles_OP();
                    actualiza_tunel_calidad();
                    configurar_limpieza();
                    CargarSemanasAnioActual();
                    ActualizarAnioReportes();
                    carga_Jefes();
                    ConfigurarTooltipParaComboBox();

                    // Configurar el ListView
                    ConfigurarListViewSimple();

                    // Cargar las semanas
                    CargarSemanasSimple();
                    //cargar años en los combos de los reportes
                    llenado_anios_concentrado_Des();
                    llenado_anios_concentrado_otras_areas();
                    llenado_anios_mermas_areas();
                    llenado_anios_mermas_Sup();
                    llenado_anios_tiempos();
                    llenado_anios_cumplimientos();

                    menuStrip1.Visible = true; // Mostrar el menú para el administrador
                }
            }
            if (nivel_user == "Administrador")
            {
                lbl_user_no_emp.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                lbl_Nom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                dtp_polvos.Value = DateTime.Now;
                dtp_tunel.Value = DateTime.Now;

                // Obtener y mostrar el número de semana inicial
                ActualizarNumeroSemana();

                actualiza_grid_users(); // Llamar al método para actualizar el DataGridView de usuarios
                actualiza_grid_Deshitratado();
                actualiza_grid_Empacado();
                actualiza_grid_inspec();
                actualiza_grid_evaporado();
                actualiza_grid_grind();
                actualiza_revolturas();
                actualiza_maquinas();
                actualiza_polvos();
                actualiza_polvos_calidad();
                actualiza_detalles_OP();
                actualiza_tunel_calidad();
                configurar_limpieza();
                CargarSemanasAnioActual();
                emaildatos();
                email_varibles();
                ActualizarAnioReportes();
                carga_Jefes();
                ConfigurarTooltipParaComboBox();
                // Configurar el ListView
                    ConfigurarListViewSimple();

                    // Cargar las semanas
                    CargarSemanasSimple();
                //cargar años en los combos de los reportes
                llenado_anios_concentrado_Des();
                llenado_anios_concentrado_otras_areas();
                llenado_anios_mermas_areas();
                llenado_anios_mermas_Sup();
                llenado_anios_tiempos();
                llenado_anios_cumplimientos();

                materialTabControl1.TabPages.Remove(tabPage3);
                materialTabControl1.TabPages.Remove(tabPage19);///eliminar pestaña de costos
                menuStrip1.Visible = false; // Mostrar el menú para el administrador
            }
            if (nivel_user == "Supervisor" || nivel_user == "Jefe de Turno")
            {
                // Ocultar las pestañas no necesarias para el usuario supervisor
                materialTabControl1.TabPages.Remove(tabPage3);
                materialTabControl1.TabPages.Remove(tabPage4);
                materialTabControl1.TabPages.Remove(tabPage11);
                materialTabControl1.TabPages.Remove(tabPage19);///eliminar pestaña de costos

                tabControl_detallesOP.TabPages.Remove(tabPage12);

                tapcontrol.TabPages.Remove(tabPage16);
                tapcontrol.TabPages.Remove(tabPage17);
                tapcontrol.TabPages.Remove(metroTabPage10);
                tapcontrol.TabPages.Remove(tabPage22);
                tapcontrol.TabPages.Remove(tabPage33);

                if(nivel_user == "Supervisor") 
                {
                    materialTabControl1.TabPages.Remove(tabPage10);
                }

                ActualizarAnioReportes();
                CargarSemanasAnioActual();
                actualiza_detalles_OP();
                email_varibles();
                carga_Jefes();
                ConfigurarTooltipParaComboBox();

                lbl_user_no_emp.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);
                lbl_Nom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point);

            }
            if (nivel_user == "Calidad")
            {
                // Ocultar las pestañas no necesarias para el usuario de calidad
                materialTabControl1.TabPages.Remove(tabPage1);
                materialTabControl1.TabPages.Remove(tabPage3);
                materialTabControl1.TabPages.Remove(tabPage4);
                materialTabControl1.TabPages.Remove(tabPage9);
                materialTabControl1.TabPages.Remove(tabPage19);///eliminar pestaña de costos

                tapcontrol.TabPages.Remove(metroTabPage9);
                tapcontrol.TabPages.Remove(tabPage16);
                tapcontrol.TabPages.Remove(tabPage17);
                tapcontrol.TabPages.Remove(tabPage18);

                dtp_polvos.Value = DateTime.Now;
                dtp_tunel.Value = DateTime.Now;
                // Obtener y mostrar el número de semana inicial
                ActualizarAnioReportes();
                CargarSemanasAnioActual();
                ActualizarNumeroSemana();
                actualiza_polvos_calidad();
                actualiza_tunel_calidad();
                configurar_limpieza();
                ConfigurarTooltipParaComboBox();

                //CARGAR COMBOS AÑOS
                llenado_anios_mermas_areas();
                llenado_anios_mermas_Sup();
                llenado_anios_tiempos();
            }
            if (nivel_user == "Mantenimiento")
            {
                // Ocultar las pestañas no necesarias para el usuario de calidad
                materialTabControl1.TabPages.Remove(tabPage1);
                materialTabControl1.TabPages.Remove(tabPage2);
                materialTabControl1.TabPages.Remove(tabPage3);
                materialTabControl1.TabPages.Remove(tabPage4);
                materialTabControl1.TabPages.Remove(tabPage9);
                materialTabControl1.TabPages.Remove(tabPage11);
                materialTabControl1.TabPages.Remove(tabPage36);
                materialTabControl1.TabPages.Remove(tabPage19);///eliminar pestaña de costos

                tapcontrol.TabPages.Remove(metroTabPage9);
                tapcontrol.TabPages.Remove(tabPage16);
                tapcontrol.TabPages.Remove(tabPage17);
                tapcontrol.TabPages.Remove(metroTabPage10);
                tapcontrol.TabPages.Remove(tabPage22);
                llenado_anios_tiempos();
            }
        }
        private void ConfigurarTooltipParaComboBox()
        {
            // Suscribirse al evento DrawItem para personalizar el dibujo
            cb_tipo_grafica.DrawMode = DrawMode.OwnerDrawFixed;

            // También puedes agregar tooltip al hover
            ToolTip toolTip = new ToolTip();
            toolTip.AutoPopDelay = 5000;
            toolTip.InitialDelay = 500;
            toolTip.ReshowDelay = 500;
            toolTip.ShowAlways = true;

            cb_tipo_grafica.MouseHover += (sender, e) =>
            {
                int index = cb_tipo_grafica.SelectedIndex;
                if (index >= 0)
                {
                    toolTip.SetToolTip(cb_tipo_grafica, cb_tipo_grafica.Items[index].ToString());
                }
            };
        }
        private void carga_Jefes()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = "SELECT \"ID_User\",\"Usuario\" FROM public.\"Usuarios\" where \"Nivel\" = 'Jefe de Turno' ORDER BY \"Usuario\" ASC;";
            dbHelper.LoadDataIntoComboBox(query, cb_jefe_turno, "Usuario", "ID_User");
        }
        private void ActualizarAnioReportes()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_grafica, "Año", null);
            CB_Anio_grafica.SelectedIndex = 0;
            CB_Anio_grafica.Focus();
            cb_Area.Focus();
        }
        private void llenado_anios_concentrado_Des()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_Con_Des, "Año", null);
            CB_Anio_Con_Des.SelectedIndex = 0;
            CB_Anio_Con_Des.Focus();
            CB_Anio_Con_Des.Focus();
        }
        private void llenado_anios_concentrado_otras_areas()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_Con_Otras_A, "Año", null);
            CB_Anio_Con_Otras_A.SelectedIndex = 0;
            CB_Anio_Con_Otras_A.Focus();
            CB_Anio_Con_Otras_A.Focus();
        }
        private void llenado_anios_mermas_areas()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_Mer_Area, "Año", null);
            CB_Anio_Mer_Area.SelectedIndex = 0;
            CB_Anio_Mer_Area.Focus();
            CB_Anio_Mer_Area.Focus();
        }
        private void llenado_anios_mermas_Sup()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_Mer_Sup, "Año", null);
            CB_Anio_Mer_Sup.SelectedIndex = 0;
            CB_Anio_Mer_Sup.Focus();
            CB_Anio_Mer_Sup.Focus();
        }
        private void llenado_anios_tiempos()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_Tiempos, "Año", null);
            CB_Anio_Tiempos.SelectedIndex = 0;
            CB_Anio_Tiempos.Focus();
            CB_Anio_Tiempos.Focus();
        }
        private void llenado_anios_cumplimientos()
        {
            // Cargar combobox
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT EXTRACT(YEAR FROM ""Fecha"") AS ""Año"" FROM public.""Ficha"" group by ""Año"" ORDER BY ""Año"" DESC ";
            dbHelper.LoadDataIntoComboBox(query, CB_Anio_Cumplimiento, "Año", null);
            CB_Anio_Cumplimiento.SelectedIndex = 0;
            CB_Anio_Cumplimiento.Focus();
            CB_Anio_Cumplimiento.Focus();
        }
        private void emaildatos() 
        {
            if (!string.IsNullOrEmpty(Tablero.Properties.Settings.Default.ServerSMTP)) 
            {
                Txt_server_Smtp.Text = Tablero.Properties.Settings.Default.ServerSMTP;
                Txt_Remitente.Text = Tablero.Properties.Settings.Default.RemitenteEMail;
                Txt_Password_SMTP.Text = Tablero.Properties.Settings.Default.PasswordEmail;
                TxtPuerto.Text = Tablero.Properties.Settings.Default.PuertoSMTP.ToString();
                Check_ssl.Checked = Tablero.Properties.Settings.Default.SSLCheck;
                Txt_Destinatarios.Text = Tablero.Properties.Settings.Default.DestinatariosEmail;
            }
        }
        private void email_varibles()
        {
            servidor_smtp = Tablero.Properties.Settings.Default.ServerSMTP;
            RemitenteEMail = Tablero.Properties.Settings.Default.RemitenteEMail;
            PasswordEmail = Tablero.Properties.Settings.Default.PasswordEmail;
            PuertoSMTP = Tablero.Properties.Settings.Default.PuertoSMTP;
            SSLCheck = Tablero.Properties.Settings.Default.SSLCheck;
            DestinatariosEmail = Tablero.Properties.Settings.Default.DestinatariosEmail;
        }
        private void CargarSemanasAnioActual()
        {
            try
            {
                // Limpiar el ComboBox primero
                cb_semana_concentrado_d.Items.Clear();
                cb_semana_concentrado_otras.Items.Clear();

                // Obtener el año actual
                int anioActual = DateTime.Now.Year;

                // Obtener el número total de semanas en el año actual
                int totalSemanas = ObtenerTotalSemanasEnAnio(anioActual);

                // Agregar las semanas al ComboBox
                for (int semana = 1; semana <= totalSemanas; semana++)
                {
                    cb_semana_concentrado_d.Items.Add(semana);
                    cb_semana_concentrado_otras.Items.Add(semana);
                    // Agregar al RadListView
                    lista_semanas.Items.Add(semana.ToString());
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al cargar semanas: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int ObtenerTotalSemanasEnAnio(int anio)
        {
            // El 28 de diciembre siempre está en la última semana del año
            DateTime ultimoDia = new DateTime(anio, 12, 28);

            // Obtener el formato de semana ISO 8601
            System.Globalization.Calendar cal = System.Globalization.CultureInfo.CurrentCulture.Calendar;
            int ultimaSemana = cal.GetWeekOfYear(ultimoDia, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            return ultimaSemana;
        }
        private void configurar_limpieza()
        {
            tabPage14.BackColor = Color.White;
            tabPage15.BackColor = Color.White;
            lbl_no_semana_polvos.BackColor = Color.White;
            lbl_no_semana_tunel.BackColor = Color.White;
            txt_no_semana_polvos.BackColor = Color.White;
            txt_no_semana_tunel.BackColor = Color.White;
            tabControl1.SelectedIndex = 0;
            tab_limpiezas.SelectedIndex = 0;
        }
        private void actualiza_grid_users()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_User"" as ""ID"", 
                            ""No_Empleado"" as ""No Empleado"", 
                            ""Usuario"" as ""Nombre de usuario"", 
                            ""Password"" as ""Contraseña"", 
                            ""Nivel""
                          FROM public.""Usuarios""
                          ORDER BY ""No_Empleado"" ASC;";  // ← ASC para orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_users, null);

            // Configurar el DataGridView
            dgv_users.Columns[0].Visible = false;
            dgv_users.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_grid_Deshitratado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""No_box_hr"" as ""Numero de Cajones/Hora"", 
                            ""Kg_fresco_hr"" as ""Kilogramos Fresco/Hora"", 
                            ""Relacion_fr_seco"" as ""Relación Fresco-Seco"",
                            ""Kg_seco_hr"" as ""Kilogramos Seco/Hora"",
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Deshidratado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_des, null);

            // Configurar el DataGridView
            dgv_metas_des.Columns[0].Visible = false;
            dgv_metas_des.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_grid_Empacado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Kg_person_hr"" as ""Kilogramos/Persona por Hora"", 
                            ""Meta_kg_hr_line"" as ""Meta Kilogramos/Hora por 1 linea"", 
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Empacado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_emp, null);

            // Configurar el DataGridView
            dgv_metas_emp.Columns[0].Visible = false;
            dgv_metas_emp.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_grid_inspec()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Kg_person_hr"" as ""Kilogramos/Persona por Hora"", 
                            ""Meta_kg_hr_line"" as ""Meta Kilogramos/Hora por 1 linea"", 
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Inspeccion""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_insp, null);

            // Configurar el DataGridView
            dgv_metas_insp.Columns[0].Visible = false;
            dgv_metas_insp.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_grid_evaporado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Capacidad_trompos"" as ""Capacidad por Trompo"", 
                            ""Cantidad_trompos"" as ""Cantidad de Trompos"",
                            ""Meta_kg_hr"" as ""Meta Kg por hora"", 
                            ""Personal_idoneo"" as ""Personal Idóneo""
                          FROM public.""Evaporado""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_Eva, null);

            // Configurar el DataGridView
            dgv_metas_Eva.Columns[0].Visible = false;
            dgv_metas_Eva.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_grid_grind()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Capacidad_Molino"" as ""Capacidad por Molino"", 
                            ""Cantidad_molinos"" as ""Cantidad de Molinos"",
                            ""Meta_Kg_hr"" as ""Meta Kg por hora"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo""
                          FROM public.""Grind""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_Grind, null);

            // Configurar el DataGridView
            dgv_metas_Grind.Columns[0].Visible = false;
            dgv_metas_Grind.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_revolturas()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Cap_Trompo_machin"" as ""Capacidad por Trompo y/o Máquinas"", 
                            ""Canti_Trompo_machin"" as ""Cantidad de Trompos y/o Máquinas"",
                            ""Meta_Kg_Hr"" as ""Meta Kg por hora"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo""
                          FROM public.""Revolturas""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_revolturas, null);

            // Configurar el DataGridView
            dgv_metas_revolturas.Columns[0].Visible = false;
            dgv_metas_revolturas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_maquinas()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Cap_Trompo_machin"" as ""Capacidad por Trompo y/o Máquinas"", 
                            ""Canti_Trompo_machin"" as ""Cantidad de Trompos y/o Máquinas"",
                            ""Meta_Kg_Hr"" as ""Meta Kg por hora"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo"",
                            ""Kg_pz"" as ""Kg por pieza""
                          FROM public.""Maquinas""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_maquinas, null);

            // Configurar el DataGridView
            dgv_metas_maquinas.Columns[0].Visible = false;
            dgv_metas_maquinas.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_polvos()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_OP"" as ""ID"", 
                            ""OP"", 
                            ""Personal_Idoneo"" as ""Personal Idóneo"", 
                            ""Meta_kg_hr_hum"" as ""Meta kg/hr Temporada Humeda"",
                            ""Meta_kg_hr_idon"" as ""Meta kg/hr Temporada Idónea"" 
                          FROM public.""Polvos""
                          ORDER BY ""OP"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_metas_polvos, null);

            // Configurar el DataGridView
            dgv_metas_polvos.Columns[0].Visible = false;
            dgv_metas_polvos.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }

        private void actualiza_detalles_OP()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta ordenada por No_Empleado de menor a mayor (ASCENDENTE)
            string querySimple = @"SELECT 
                            ""ID_Dt_OP"" as ""ID"", 
                            ""Orden_Produccion"" as ""Orden de Producción"", 
                            ""Producto"", 
                            ""Medida"",
                            ""Descripcion"" as ""Descripción"", 
                            ""Especificacion"" as ""Especificación"",
                            ""Ingredientes"",
                            ""Humedad"",
                            ""Comercio"",
                            ""Manzana"",
                            ""Analisis"",
                            ""Area_Proceso"" as ""Área de Proceso"",
                            ""OP_Origen"" as ""OP de Origen"",
                            ""Destino"",
                            ""Clasificacion"" as ""Clasificación""
                          FROM public.""Detalles_OP""
                          ORDER BY ""Orden_Produccion"" ASC;";  // ← orden ascendente

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_detalles_op, null);

            // Configurar el DataGridView
            dgv_detalles_op.Columns[0].Visible = false;
            dgv_detalles_op.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
        }
        private void actualiza_polvos_calidad()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta para la tabla Ficha filtrando por Área = 'Polvos'
            string querySimple = @"SELECT 
                    ""ID_Limpieza"" as ""ID"", 
                    ""Fecha"", 
                    EXTRACT(WEEK FROM ""Fecha"") as ""No Semana"",
                    ""Kg_merma"" as ""Kg de merma por Limpieza""
                  FROM public.""Limpieza_polvos""
                  ORDER BY ""Fecha"" DESC;";

            // Cargar los datos de la tabla Ficha en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_polvos_calidad, null);

            // Configurar el DataGridView
            dgv_polvos_calidad.Columns[0].Visible = false;

            // Configurar formato para la columna de fecha
            if (dgv_polvos_calidad.Columns["Fecha"] != null)
            {
                dgv_polvos_calidad.Columns["Fecha"].DefaultCellStyle.Format = "dd/MM/yyyy";
            }
        }
        private void actualiza_tunel_calidad()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Consulta para la tabla Ficha filtrando por Área = 'Polvos'
            string querySimple = @"SELECT 
                    ""ID_Tunel"" as ""ID"", 
                    ""Fecha"", 
                    EXTRACT(WEEK FROM ""Fecha"") as ""No Semana"",
                    ""Kg_merma"" as ""Kg de merma de Tunel""
                  FROM public.""Limpieza_tunel""
                  ORDER BY ""Fecha"" DESC;";

            // Cargar los datos de la tabla Ficha en el DataGridView
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_Tunel_calidad, null);

            // Configurar el DataGridView
            dgv_Tunel_calidad.Columns[0].Visible = false;

            // Configurar formato para la columna de fecha
            if (dgv_Tunel_calidad.Columns["Fecha"] != null)
            {
                dgv_Tunel_calidad.Columns["Fecha"].DefaultCellStyle.Format = "dd/MM/yyyy";
            }
        }
        private void actualiza_reporte_merma()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Obtener el año seleccionado del ComboBox
            string añoSeleccionado = CB_Anio_Mer_Area.Text;

            // Validar que se haya seleccionado un año
            if (string.IsNullOrEmpty(añoSeleccionado))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Por favor, seleccione un año antes de generar el reporte de mermas.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string querySimple = @"
WITH semanas AS (
    SELECT 
        ""Area"", 
        EXTRACT(WEEK FROM ""Fecha"") AS numero_semana, 
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        
        -- Cálculo para Tunel/Sumergidor
        CASE 
            WHEN ""Area"" = 'Tunel/Sumergidor' THEN
                SUM(""Merma_podrido"" + ""Merma_tina"" + ""Merma_piso"" + ""Merma_canaletas"" + ""Merma_lavado_bandas"") / 
                NULLIF(SUM(""Kg_enter_proceso"" - ""Merma_canica""), 0)
        END AS merma_tunel,

        -- Cálculo para Despegue (se unirá con Limpieza_tunel)
        CASE 
            WHEN ""Area"" = 'Despegue' THEN
                SUM(""Merma_kg"") / NULLIF(SUM(""Kg_prod_seco""), 0)
        END AS merma_despegue_base,

        -- Cálculo para otras áreas
        CASE 
            WHEN ""Area"" IN ('Evaporado', 'Grind', 'Inspección', 'Empacado', 'Revolturas', 'Máquinas') THEN
                SUM(""Merma_kg"") / NULLIF(SUM(""Kg_enter_proceso""), 0)
        END AS merma_otras_areas,

        -- Cálculo para Polvos (se unirá con Limpieza_polvos)
        CASE 
            WHEN ""Area"" = 'Polvos' THEN
                SUM(""Merma_kg"") / NULLIF(SUM(""Kg_prod_seco""), 0)
        END AS merma_polvos_base,

        -- Sumas para Despegue
        SUM(CASE WHEN ""Area"" = 'Despegue' THEN ""Merma_kg"" ELSE 0 END) AS suma_merma_despegue,
        SUM(CASE WHEN ""Area"" = 'Despegue' THEN ""Kg_prod_seco"" ELSE 0 END) AS suma_kg_seco_despegue,

        -- Sumas para Polvos
        SUM(CASE WHEN ""Area"" = 'Polvos' THEN ""Merma_kg"" ELSE 0 END) AS suma_merma_polvos,
        SUM(CASE WHEN ""Area"" = 'Polvos' THEN ""Kg_prod_seco"" ELSE 0 END) AS suma_kg_seco_polvos

    FROM public.""Ficha""
    WHERE EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY ""Area"", EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
),

limpieza_tunel_semanal AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS numero_semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        SUM(""Kg_merma"") AS kg_merma_tunel_semana
    FROM public.""Limpieza_tunel""
    WHERE EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
),

limpieza_polvos_semanal AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS numero_semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        SUM(""Kg_merma"") AS kg_merma_polvos_semana
    FROM public.""Limpieza_polvos""
    WHERE EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
)

SELECT 
    s.""Area"",
    s.numero_semana as ""Numero de Semana"",
    s.año as ""Año"",
    CASE 
        WHEN s.""Area"" = 'Tunel/Sumergidor' THEN
            ROUND(COALESCE(s.merma_tunel, 0) * 100, 2)

        WHEN s.""Area"" = 'Despegue' THEN
            ROUND(
                COALESCE(
                    (s.suma_merma_despegue + COALESCE(lt.kg_merma_tunel_semana, 0)) / 
                    NULLIF(s.suma_kg_seco_despegue, 0), 0) * 100, 2)

        WHEN s.""Area"" IN ('Evaporado', 'Grind', 'Inspección', 'Empacado', 'Revolturas', 'Máquinas') THEN
            ROUND(COALESCE(s.merma_otras_areas, 0) * 100, 2)

        WHEN s.""Area"" = 'Polvos' THEN
            ROUND(COALESCE((s.suma_merma_polvos + COALESCE(lp.kg_merma_polvos_semana, 0)) / NULLIF(s.suma_kg_seco_polvos, 0), 0) * 100, 2)
        ELSE 0
    END AS ""Porcentaje de Merma""

FROM semanas s
LEFT JOIN limpieza_tunel_semanal lt ON s.numero_semana = lt.numero_semana AND s.año = lt.año
LEFT JOIN limpieza_polvos_semanal lp ON s.numero_semana = lp.numero_semana AND s.año = lp.año
WHERE s.""Area"" IS NOT NULL
ORDER BY s.año, s.numero_semana, s.""Area"";";

            // Crear el parámetro para el año
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@Anio", NpgsqlTypes.NpgsqlDbType.Integer)
        {
            Value = Convert.ToInt32(añoSeleccionado)
        }
            };

            // Cargar los datos de la tabla Ficha en el DataGridView con el filtro de año
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_reporte_merma, parameters);
        }
        private void actualiza_reporte_merma_S()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Obtener el año seleccionado del ComboBox
            string añoSeleccionado = CB_Anio_Mer_Sup.Text;

            // Validar que se haya seleccionado un año
            if (string.IsNullOrEmpty(añoSeleccionado))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Por favor, seleccione un año antes de generar el reporte de mermas por supervisor.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Consulta para la tabla Ficha filtrando por Área = 'Polvos'
            string querySimple = @"
WITH merma_usuario AS (
    SELECT 
        u.""Usuario"" as ""Nombre_Usuario"",
        EXTRACT(WEEK FROM f.""Fecha"") AS numero_semana, 
        EXTRACT(YEAR FROM f.""Fecha"") AS año,
        f.""Area"" As Area,
        SUM(
            CASE 
                WHEN f.""Area"" = 'Tunel/Sumergidor' THEN
                    f.""Merma_podrido"" + f.""Merma_tina"" + f.""Merma_piso"" + f.""Merma_canaletas"" + f.""Merma_lavado_bandas""
                ELSE f.""Merma_kg""
            END
        ) AS total_merma_kg
    FROM public.""Ficha"" f
    INNER JOIN public.""Usuarios"" u ON f.""ID_user"" = u.""ID_User""
    WHERE f.""Area"" IS NOT NULL
        AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio
    GROUP BY u.""Usuario"", EXTRACT(WEEK FROM f.""Fecha""), EXTRACT(YEAR FROM f.""Fecha""), f.""Area""
)

SELECT 
    ""Nombre_Usuario"" AS ""Supervisor"",
    numero_semana as ""Numero de Semana"",
    año as ""Año"",
    Area as ""Área"",
    COALESCE(total_merma_kg, 0) as ""Merma(Kg)""
FROM merma_usuario
ORDER BY año, numero_semana, ""Nombre_Usuario"";";

            // Crear el parámetro para el año
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@Anio", NpgsqlTypes.NpgsqlDbType.Integer)
        {
            Value = Convert.ToInt32(añoSeleccionado)
        }
            };

            // Cargar los datos de la tabla Ficha en el DataGridView con el filtro de año
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_reporte_merma_S, parameters);
        }
        private void cb_Area_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                //TUNEL/ SUMERGIDOR
                if (!editar) { reiniciarCampos(); }

                //hacer visible para tunel
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;

                //habilitar controles
                cb_Turno.Enabled = true;
                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;


                //nombrar controles
                Txt_1.EmbeddedLabelText = "Lote";
                Txt_2.EmbeddedLabelText = "Kg Entrada (Proceso)";
                Txt_3.EmbeddedLabelText = "Canica";
                Txt_4.EmbeddedLabelText = "Merma Podrido";
                Txt_5.EmbeddedLabelText = "Merma de Tina";
                Txt_6.EmbeddedLabelText = "Merma de Piso";
                Txt_7.EmbeddedLabelText = "Merma Canaletas";
                Txt_8.EmbeddedLabelText = "Merma Lavado de Bandas";
                Txt_9.EmbeddedLabelText = "Personal Operativo";
                Txt_10.EmbeddedLabelText = "Cascara y Carrete";

                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta Programada";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Kg Frescos de Entrada a secador";

                //hacer visibles controles
                Txt_1.Visible = true;
                Txt_6.Visible = true;
                Txt_7.Visible = true;
                Txt_8.Visible = true;
                Txt_9.Visible = true;
                Txt_10.Visible = true;

                //hacer invisibles controles

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                Txt_11.Visible = false;
                radMultiColumnComboBox1.Visible = false;
                cb_proceso.Visible = false;

                //Inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;
                Txt_7.Enabled = false;
                Txt_8.Enabled = false;
                Txt_9.Enabled = false;
                Txt_10.Enabled = false;
                Txt_11.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                Txt_Read_1.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Deshidratado\" ORDER BY \"OP\";";
                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 1)
            {
                //Despegue
                if (!editar) { reiniciarCampos(); }

                //hacer visible para Despegue
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;

                //Txt_1.Visible = false;

                //habilitar controles
                cb_Turno.Enabled = false;
                cb_OP.Enabled = false;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                if (editar)
                {
                    Txt_1.EmbeddedLabelText = "Lote";
                    Txt_1.Visible = true;
                    Txt_1.Enabled = false;
                    radMultiColumnComboBox1.Visible = false;
                }
                else
                {
                    Txt_1.Visible = false;
                    radMultiColumnComboBox1.Visible = true;
                    radMultiColumnComboBox1.Enabled = true;
                }
                Txt_1.EmbeddedLabelText = "Lote";
                Txt_2.EmbeddedLabelText = "Kilos Producto Seco";
                Txt_3.EmbeddedLabelText = "Merma Kg Secos";
                Txt_4.EmbeddedLabelText = "Kg Secos Fuera de Espec.";
                Txt_5.EmbeddedLabelText = "Kg para Resecar";
                Txt_6.EmbeddedLabelText = "Personal Operativo";

                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta Programada";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Kg Frescos Entrada a secador";
                Txt_Read_5.EmbeddedLabelText = "Cumplimiento de Metas (%)";
                Txt_Read_7.EmbeddedLabelText = "Relación Fresco-Seco";
                Txt_Read_8.EmbeddedLabelText = "FTT";

                Txt_6.Visible = true;
                Txt_Read_5.Visible = true;
                Txt_Read_7.Visible = true;
                Txt_Read_8.Visible = true;

                Txt_Read_9.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;
                cb_proceso.Visible = false;
                Txt_Read_6.Visible = false;

                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;

                //Inabilitar controles
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;
                Txt_Read_1.Enabled = false;
                Txt_Read_2.Enabled = false;
                Txt_Read_3.Enabled = false;
                Txt_Read_4.Enabled = false;
                Txt_Read_5.Enabled = false;
                Txt_Read_6.Enabled = false;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Deshidratado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");

                // Cargar RadMultiColumnComboBox con Lote y OP

                query = @"SELECT ""ID_Ficha"", ""Lote"", ""OP"" 
                          FROM public.""Ficha"" 
                          WHERE ""Terminado_Tunel"" = false 
                            AND ""Area"" = 'Tunel/Sumergidor' 
                          ORDER BY ""Lote"";";

                string[] columnas = { "ID_Ficha", "Lote", "OP" };
                string[] headers = { "ID", "Lote", "OP" };
                bool[] visibilidad = { false, true, true };

                dbHelper.LoadDataIntoMultiColumnComboBox(
                    query,
                    radMultiColumnComboBox1,
                    columnas,
                    headers,
                    visibilidad
                );
                // DESPUÉS de cargar, establecer SelectedIndex = -1
                radMultiColumnComboBox1.SelectedIndex = -1;
                radMultiColumnComboBox1.Text = ""; // Limpiar el texto mostrado
            }
            if (cb_Area.SelectedIndex == 2)
            {
                //Evaporado
                reiniciarCampos();

                radMultiColumnComboBox1.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                Txt_Read_5.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";
                Txt_Read_5.EmbeddedLabelText = "Aumento por humedad (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Evaporado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 3)
            {
                //Grind
                reiniciarCampos();

                radMultiColumnComboBox1.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Grind\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 4)
            {
                //Inspeccion
                reiniciarCampos();

                radMultiColumnComboBox1.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                cb_proceso.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;
                cb_proceso.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Inspeccion\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 5)
            {
                //Empacado
                reiniciarCampos();

                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;
                cb_proceso.Visible = false;
                radMultiColumnComboBox1.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Empacado\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 6)
            {
                //Polvos
                reiniciarCampos();

                radMultiColumnComboBox1.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                Txt_6.Visible = true;
                Txt_7.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma PT";
                Txt_5.EmbeddedLabelText = "Personal Operativo";
                Txt_6.EmbeddedLabelText = "Polvo de colector";
                Txt_7.EmbeddedLabelText = "Granulo";

                //hacer invisibles controles
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;
                Txt_7.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Polvos\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 7)
            {
                //Revolturas
                reiniciarCampos();

                radMultiColumnComboBox1.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Kg entrada (proceso)";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";

                //hacer invisibles controles
                Txt_6.Visible = false;
                Txt_7.Visible = false;
                Txt_8.Visible = false;
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;

                Txt_Read_5.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;
                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Revolturas\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
            if (cb_Area.SelectedIndex == 8)
            {
                //Máquinas
                reiniciarCampos();

                radMultiColumnComboBox1.Visible = false;
                card_datos.Visible = true;
                card_TM.Visible = true;
                card_botones.Visible = true;
                card_meal_energy.Visible = true;
                Txt_1.Visible = true;
                Txt_6.Visible = true;
                Txt_7.Visible = true;
                Txt_8.Visible = true;

                Txt_1.Text = string.Empty;

                //habilitar controles

                cb_OP.Enabled = true;
                dtp1.Enabled = true;
                cb_jefe_turno.Enabled = true;

                //nombrar controles
                Txt_1.EmbeddedLabelText = "Piezas Producidas";
                Txt_2.EmbeddedLabelText = "Kg Producto Terminado";
                Txt_3.EmbeddedLabelText = "Kg fuera de Espec.";
                Txt_4.EmbeddedLabelText = "Merma en Kg";
                Txt_5.EmbeddedLabelText = "Personal Operativo";
                Txt_6.EmbeddedLabelText = "Bobina Kg Entrada";
                Txt_7.EmbeddedLabelText = "Bobina Utilizada";
                Txt_8.EmbeddedLabelText = "Bobina Merma";

                //nombrar controles
                Txt_Read_1.EmbeddedLabelText = "Horas Progamadas";
                Txt_Read_2.EmbeddedLabelText = "Meta en Kg";
                Txt_Read_3.EmbeddedLabelText = "Horas Efectivas";
                Txt_Read_4.EmbeddedLabelText = "Logro de Planeación (%)";
                Txt_Read_5.EmbeddedLabelText = "Kg entrada (proceso)";

                //hacer invisibles controles
                Txt_9.Visible = false;
                Txt_10.Visible = false;
                Txt_11.Visible = false;
                Txt_Read_6.Visible = false;
                Txt_Read_7.Visible = false;
                Txt_Read_8.Visible = false;

                //inabilitar controles
                cb_Turno.Enabled = false;
                Txt_1.Enabled = false;
                Txt_2.Enabled = false;
                Txt_3.Enabled = false;
                Txt_4.Enabled = false;
                Txt_5.Enabled = false;
                Txt_6.Enabled = false;
                Txt_7.Enabled = false;
                Txt_8.Enabled = false;
                Mask_txt_hr1.Enabled = false;
                Mask_txt_hr2.Enabled = false;
                txt_Tiempo_comida.Enabled = false;
                txt_Tiempo_energia.Enabled = false;
                dgv_mecanico.Enabled = false;
                dgv_operativo.Enabled = false;
                cb_proceso.Enabled = false;

                Txt_Read_9.Visible = false;
                cb_proceso.Visible = false;

                // Cargar combobox OP
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = "SELECT \"ID_OP\", \"OP\" FROM public.\"Maquinas\" ORDER BY \"OP\";";

                dbHelper.LoadDataIntoComboBox(query, cb_OP, "OP", "ID_OP");
            }
        }
        private void CalcularSuma()
        {
            decimal total = 0;
            decimal Kg_entrada = 0;
            // Lista de los TextBox que quieres sumar
            RadTextBox[] cajas = { Txt_3, Txt_4, Txt_5, Txt_6, Txt_7, Txt_8 };

            foreach (RadTextBox txt in cajas)
            {
                if (decimal.TryParse(txt.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal valor) ||
                    decimal.TryParse(txt.Text, NumberStyles.Number, CultureInfo.CurrentCulture, out valor))
                {
                    total += valor;
                }
            }

            // Para Txt_2 usar TryParse (evita excepciones mientras el usuario escribe)
            if (!decimal.TryParse(Txt_2.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out Kg_entrada) &&
                !decimal.TryParse(Txt_2.Text, NumberStyles.Number, CultureInfo.CurrentCulture, out Kg_entrada))
            {
                Kg_entrada = 0m;
            }

            decimal resultado = Kg_entrada - total;
            Txt_Read_4.Text = resultado.ToString("0.00", CultureInfo.CurrentCulture);
        }

        private void reiniciarCampos()
        {
            cb_Area.Enabled = true;
            cb_Turno.SelectedIndex = -1;
            cb_Turno.Focus();
            cb_OP.Enabled = true;
            cb_OP.SelectedIndex = -1;
            cb_OP.Focus();
            cb_proceso.SelectedIndex = -1;
            cb_proceso.Focus();
            cb_jefe_turno.SelectedIndex = -1;
            cb_jefe_turno.Focus();
            cb_jefe_turno.Enabled = false;
            dtp1.Value = DateTime.Now;
            Txt_1.Text = string.Empty;
            Txt_2.Text = string.Empty;
            Txt_3.Text = string.Empty;
            Txt_4.Text = string.Empty;
            Txt_5.Text = string.Empty;
            Txt_6.Text = string.Empty;
            Txt_7.Text = string.Empty;
            Txt_8.Text = string.Empty;
            Txt_9.Text = string.Empty;
            Txt_10.Text = string.Empty;
            Txt_11.Text = string.Empty;
            Txt_meta.Text = string.Empty;
            Txt_Read_1.Text = string.Empty;
            Txt_Read_2.Text = string.Empty;
            Txt_Read_3.Text = string.Empty;
            Txt_Read_4.Text = string.Empty;
            Txt_Read_5.Text = string.Empty;
            Txt_Read_6.Text = string.Empty;
            Txt_Read_7.Text = string.Empty;
            Txt_Read_8.Text = string.Empty;
            Txt_Read_9.Text = string.Empty;
            txt_Tiempo_comida.Text = "30";
            //hacer invisibles controles
            card_datos.Visible = false;
            card_TM.Visible = false;
            card_botones.Visible = false;
            card_meal_energy.Visible = false;
            //deshabilitar controles
            cb_Turno.Enabled = false;
            cb_OP.Enabled = false;
            dtp1.Enabled = false;
            txt_Tiempo_energia.Text = "0";
            if (dgv_mecanico.DataSource is DataTable dt1)
            {
                dt1.Rows.Clear();
            }else
            {
                dgv_mecanico.Rows.Clear();
            }

            if (dgv_operativo.DataSource is DataTable dt2)
            {
                dt2.Rows.Clear();
            }
            else 
            {                 
                dgv_operativo.Rows.Clear(); 
            }
                
        }
        private void reiniciarCampos_comboturno()
        {
            Mask_txt_hr1.Text = string.Empty;
            Mask_txt_hr2.Text = string.Empty;
            txt_Tiempo_comida.Text = "30";
            txt_Tiempo_energia.Text = "0";
            if (dgv_mecanico.DataSource is DataTable dt1)
            {
                dt1.Rows.Clear();
            }
            else
            {
                dgv_mecanico.Rows.Clear();
            }

            if (dgv_operativo.DataSource is DataTable dt2)
            {
                dt2.Rows.Clear();
            }
            else
            {
                dgv_operativo.Rows.Clear();
            }

        }


        private void dgv_users_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Obtener el valor de la primera columna del renglón clicado
                id_global_users = dgv_users.Rows[e.RowIndex].Cells[0].Value.ToString();
                // Cargar los datos del usuario seleccionado en los campos de texto
                txt_no_emp.Text = dgv_users.Rows[e.RowIndex].Cells[1].Value.ToString(); // No Empleado
                txt_usuario.Text = dgv_users.Rows[e.RowIndex].Cells[2].Value.ToString(); // Nombre de usuario
                txt_contra.Text = dgv_users.Rows[e.RowIndex].Cells[3].Value.ToString(); // Contraseña
                cmb_nivel_user.SelectedItem = dgv_users.Rows[e.RowIndex].Cells[4].Value.ToString(); // Nivel

                btn_edit.Enabled = true;
                btn_delete_user.Enabled = true;
                cmb_nivel_user.Enabled = true;
                cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
                cmb_nivel_user.Enabled = false;
            }
        }
        private void LimpiarComboBox(MaterialComboBox comboBox)
        {
            comboBox.DataSource = null;
            comboBox.Items.Clear();
            comboBox.Text = string.Empty;
        }
        private void btn_save_Click(object sender, EventArgs e)
        {
            if (txt_no_emp.Text == string.Empty || txt_usuario.Text == string.Empty || txt_contra.Text == string.Empty || cmb_nivel_user.SelectedIndex == -1)
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_users))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idUsuario = System.Convert.ToInt32(id_global_users);

                    queryInsertUpdate = "UPDATE public.\"Usuarios\" SET \"Usuario\" = @Usuario, \"Password\" = @Password, \"Nivel\" = @Nivel  WHERE \"ID_User\" = @ID_usuario;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Usuario", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_usuario.Text
                        },
                        new NpgsqlParameter("@Password", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_contra.Text
                        },
                        new NpgsqlParameter("@Nivel", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = cmb_nivel_user.SelectedItem?.ToString() ?? ""
                        },
                        new NpgsqlParameter("@ID_usuario", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idUsuario  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                    LimpiarComboBox(cb_jefe_turno);
                    carga_Jefes();
                }
                else
                {

                    // Verificar si el usuario ya existe
                    string queryCheckUser = "SELECT COUNT(*) FROM public.\"Usuarios\" WHERE \"Usuario\"  ILIKE @Usuario or \"No_Empleado\" ILIKE @no_emp;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@Usuario", txt_usuario.Text),
                    new NpgsqlParameter("@no_emp", txt_no_emp.Text)
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckUser, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El nombre de usuario y/o numero de emplado ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Usuarios\" (\"No_Empleado\", \"Usuario\", \"Password\", \"Nivel\") VALUES (@No_Empleado, @Usuario, @Password, @Nivel);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@No_Empleado", txt_no_emp.Text),
                        new NpgsqlParameter("@Usuario", txt_usuario.Text),
                        new NpgsqlParameter("@Password", txt_contra.Text),
                        new NpgsqlParameter("@Nivel", cmb_nivel_user.SelectedItem.ToString())
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                    LimpiarComboBox(cb_jefe_turno);
                    carga_Jefes();
                }

                if (result > 0)
                {
                    actualiza_grid_users(); // Actualizar el DataGridView de usuarios
                    limpiarCampos(); // Limpiar los campos de texto
                    btn_save.Enabled = false;
                    btn_cancel.Enabled = false;
                    txt_no_emp.Enabled = false;
                    txt_usuario.Enabled = false;
                    txt_contra.Enabled = false;
                    btn_edit.Enabled = false;
                    id_global_users = string.Empty;

                    cmb_nivel_user.Enabled = true;
                    cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
                    cmb_nivel_user.Enabled = false;
                }
            }
        }
        private void limpiarCampos()
        {
            txt_no_emp.Text = string.Empty;
            txt_usuario.Text = string.Empty;
            txt_contra.Text = string.Empty;
            cmb_nivel_user.SelectedIndex = -1;
            id_global_users = string.Empty; // Limpiar la variable global
            cb_Area.SelectedIndex = -1; // Limpiar el ComboBox de áreas
            cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
            txt_no_emp.Focus(); // Enfocar el campo de No Empleado
        }
        private void limpiarCampos_meta()
        {
            cmb_area.Enabled = true;
            cmb_area.SelectedIndex = -1;
            cmb_area.Focus();
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;
            txt_op.Focus();
        }
        private void limpiarCampos_detalles_OP()
        {
            txt_orden_produc.Text = string.Empty;
            txt_producto.Text = string.Empty;
            txt_medida.Text = string.Empty;
            txt_descripcion.Text = string.Empty;
            txt_especificacion.Text = string.Empty;
            txt_ingredientes.Text = string.Empty;
            txt_humedad.Text = string.Empty;
            cb_Comercio.SelectedIndex = -1;
            cb_Manzana.SelectedIndex = -1;
            txt_analisis.Text = string.Empty;
            txt_area_proceso.Text = string.Empty;
            txt_op_origen.Text = string.Empty;
            txt_destino.Text = string.Empty;
            txt_clacificacion.Text = string.Empty;
            txt_orden_produc.Focus();
        }
        private void btn_new_user_Click(object sender, EventArgs e)
        {
            // Obtener el valor máximo de la columna "No Empleado" (índice 1)
            int maxNoEmp = 0;
            foreach (DataGridViewRow row in dgv_users.Rows)
            {
                if (row.IsNewRow) continue; // Ignorar la fila para agregar nuevo
                int value;
                if (int.TryParse(row.Cells[1].Value?.ToString(), out value))
                {
                    if (value > maxNoEmp)
                        maxNoEmp = value;
                }
            }
            txt_no_emp.Text = (maxNoEmp + 1).ToString();

            btn_save.Enabled = true;
            btn_cancel.Enabled = true;
            btn_edit.Enabled = false;
            btn_delete_user.Enabled = false;

            txt_no_emp.Enabled = true;
            txt_usuario.Enabled = true;
            txt_contra.Enabled = true;
            cmb_nivel_user.Enabled = true;


            txt_usuario.Text = string.Empty;
            txt_contra.Text = string.Empty;
            cmb_nivel_user.SelectedIndex = -1;
            id_global_users = string.Empty; // Limpiar la variable global
            cb_Area.SelectedIndex = -1; // Limpiar el ComboBox de áreas
            cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
            txt_usuario.Focus(); // Enfocar el campo de No Empleado
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = false;
            btn_cancel.Enabled = false;
            btn_delete_user.Enabled = true;
            limpiarCampos();
            txt_no_emp.Enabled = false;
            txt_usuario.Enabled = false;
            txt_contra.Enabled = false;
            btn_edit.Enabled = false;
            id_global_users = string.Empty;

            cmb_nivel_user.Enabled = true;
            cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
            cmb_nivel_user.Enabled = false;
        }

        private void btn_edit_Click(object sender, EventArgs e)
        {
            btn_save.Enabled = true;
            btn_cancel.Enabled = true;
            btn_delete_user.Enabled = false;
            txt_usuario.Enabled = true;
            txt_contra.Enabled = true;
            cmb_nivel_user.Enabled = true;
            cmb_nivel_user.Focus();
            txt_usuario.Focus();
        }

        private void txt_no_emp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void btn_search_user_Click(object sender, EventArgs e)
        {
            // Si el filtro está activo, limpiar filtro y campos de búsqueda
            if (filtroUsuariosActivo)
            {
                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_users.Rows)
                {
                    row.Visible = true;
                }

                // Restablecer el estado del filtro
                filtroUsuariosActivo = false;
            }

            string noEmp = txt_search_no_emp.Text.Trim();
            string nomEmp = txt_search_nom_emp.Text.Trim();
            string nivel = cmb_serch_nivel.Text.Trim();

            // Validar que al menos un campo esté lleno
            if (string.IsNullOrEmpty(noEmp) && string.IsNullOrEmpty(nomEmp) && string.IsNullOrEmpty(nivel))
            {
                MetroFramework.MetroMessageBox.Show(this, "Favor de llenar uno o más campos para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Deseleccionar cualquier celda y fila antes de filtrar
            dgv_users.ClearSelection();
            dgv_users.CurrentCell = null;

            // Mover el CurrencyManager a una posición válida
            CurrencyManager cm = (CurrencyManager)BindingContext[dgv_users.DataSource];
            if (cm != null && cm.Count > 0)
                cm.Position = -1;

            bool hayFiltro = false;
            foreach (DataGridViewRow row in dgv_users.Rows)
            {
                if (row.IsNewRow) continue;

                bool match = true;

                if (!string.IsNullOrEmpty(noEmp))
                {
                    match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(noEmp, StringComparison.OrdinalIgnoreCase);
                }
                if (!string.IsNullOrEmpty(nomEmp))
                {
                    match &= row.Cells[2].Value != null && row.Cells[2].Value.ToString().IndexOf(nomEmp, StringComparison.OrdinalIgnoreCase) >= 0;
                }
                if (!string.IsNullOrEmpty(nivel))
                {
                    match &= row.Cells[4].Value != null && row.Cells[4].Value.ToString().Equals(nivel, StringComparison.OrdinalIgnoreCase);
                }

                row.Visible = match;
                if (match) hayFiltro = true;
            }
            filtroUsuariosActivo = hayFiltro;

            if (!hayFiltro)
            {
                MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                    "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_search_limpiar_Click(object sender, EventArgs e)
        {
            limpiar_filtros();
        }
        private void limpiar_filtros()
        {
            // Mostrar todas las filas del DataGridView
            foreach (DataGridViewRow row in dgv_users.Rows)
            {
                row.Visible = true;
            }

            // Limpiar los campos de búsqueda
            txt_search_no_emp.Text = string.Empty;
            txt_search_nom_emp.Text = string.Empty;
            cmb_serch_nivel.SelectedIndex = -1;
            cmb_serch_nivel.Focus();
            txt_search_no_emp.Focus();

            // Restablecer el estado del filtro
            filtroUsuariosActivo = false;
        }
        private void limpiar_filtros_OP()
        {
            // Mostrar todas las filas del DataGridView
            // Mostrar todas las filas del DataGridView Deshidratado
            if (filtroUsuariosActivo_OP_dgv_activo == 1)
            {
                foreach (DataGridViewRow row in dgv_metas_des.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Empacado
            if (filtroUsuariosActivo_OP_dgv_activo == 2)
            {
                foreach (DataGridViewRow row in dgv_metas_emp.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Inspección
            if (filtroUsuariosActivo_OP_dgv_activo == 3)
            {
                foreach (DataGridViewRow row in dgv_metas_insp.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Evaporado
            if (filtroUsuariosActivo_OP_dgv_activo == 4)
            {
                foreach (DataGridViewRow row in dgv_metas_Eva.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Grind
            if (filtroUsuariosActivo_OP_dgv_activo == 5)
            {
                foreach (DataGridViewRow row in dgv_metas_Grind.Rows)
                {
                    row.Visible = true;
                }
            }
            // Mostrar todas las filas del DataGridView Revolturas
            if (filtroUsuariosActivo_OP_dgv_activo == 6)
            {
                foreach (DataGridViewRow row in dgv_metas_revolturas.Rows)
                {
                    row.Visible = true;
                }
            }

            // Limpiar los campos de búsqueda
            txt_OP_Search.Text = string.Empty;
            cmb_area_search.SelectedIndex = -1;
            cmb_area_search.Focus();

            txt_OP_Search.Enabled = false;

            // Restablecer el estado del filtro
            filtroUsuariosActivo_OP = false;
            filtroUsuariosActivo_OP_dgv_activo = 0;
        }
        private void txt_search_no_emp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void cmb_area_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;
            txt_op.Focus();
            reiniciar_controles_meta();
            if (cmb_area.SelectedIndex == 0)
            {
                //habilitar controles
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_5.Enabled = true;

                txt_Meta_4.Visible = true;
                txt_Meta_5.Visible = true;

                //nombrar controles
                txt_Meta_1.Hint = "No. Cajones por Hora";
                txt_Meta_2.Hint = "Kg Fresco por Hora";
                txt_Meta_3.Hint = "Relación Fresco-Seco";
                txt_Meta_4.Hint = "Kg Seco/Hr";
                txt_Meta_5.Hint = "Personal Idóneo";

                tap_control_metas.SelectedIndex = 0;
            }
            if (cmb_area.SelectedIndex == 1)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Kg por persona por Hora";
                txt_Meta_3.Hint = "Meta Kg por hr por 1 linea";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;

                //hacer invisibles controles
                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 1;
            }
            if (cmb_area.SelectedIndex == 2)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Kg por persona por Hora";
                txt_Meta_3.Hint = "Meta Kg/hr por 1 linea";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;

                //hacer invisibles controles
                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 2;
            }
            if (cmb_area.SelectedIndex == 3)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad por trompo";
                txt_Meta_3.Hint = "Cantidad de trompos";
                txt_Meta_4.Hint = "Meta Kg por hora";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Visible = true;

                //hacer invisibles controles
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 3;
            }
            if (cmb_area.SelectedIndex == 4)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad por molino";
                txt_Meta_3.Hint = "Cantidad de molinos";
                txt_Meta_4.Hint = "Meta Kg por hora";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Visible = true;

                //hacer invisibles controles
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 4;
            }
            if (cmb_area.SelectedIndex == 5)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad/(Trompo y/o Máq.)";
                txt_Meta_3.Hint = "Cantidad/(Trompo y/o Máq.)";
                txt_Meta_4.Hint = "Meta Kg por hora";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_4.Visible = true;

                //hacer invisibles controles
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 5;
            }
            if (cmb_area.SelectedIndex == 6)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Meta kg/hr temporada humenda";
                txt_Meta_3.Hint = "Meta kg/hr temporada idónea";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;

                //hacer invisibles controles
                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                tap_control_metas.SelectedIndex = 6;
            }
            if (cmb_area.SelectedIndex == 7)
            {
                //nombrar controles
                txt_op.Enabled = true;
                txt_Meta_1.Hint = "Personal Idóneo";
                txt_Meta_2.Hint = "Capacidad/(Trompo y/o Máq.)";
                txt_Meta_3.Hint = "Cantidad/(Trompo y/o Máq.)";
                txt_Meta_4.Hint = "Meta Kg por hora";
                txt_Meta_5.Hint = "Kg por pieza";

                //habilitar controles
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_5.Enabled = true;

                //hacer visible controles
                txt_Meta_4.Visible = true;
                txt_Meta_5.Visible = true;

                tap_control_metas.SelectedIndex = 7;
            }
        }
        private void reiniciar_controles_meta()
        {
            txt_Meta_1.Hint = string.Empty;
            txt_Meta_2.Hint = string.Empty;
            txt_Meta_3.Hint = string.Empty;
            txt_Meta_4.Hint = string.Empty;
            txt_Meta_5.Hint = string.Empty;
            //deshabilitar controles
            txt_op.Enabled = false;
            txt_Meta_1.Enabled = false;
            txt_Meta_2.Enabled = false;
            txt_Meta_3.Enabled = false;
            txt_Meta_4.Enabled = false;
            txt_Meta_5.Enabled = false;
        }

        private void btn_new_op_Click(object sender, EventArgs e)
        {
            //habilitar controles
            cmb_area.Enabled = true;
            btn_meta_cancel.Enabled = true;
            btn_meta_save.Enabled = true;
            btn_meta_edit.Enabled = false;
            btn_meta_delete.Enabled = false;

            //limpiar campos
            cmb_area.SelectedIndex = -1;
            txt_op.Text = string.Empty;
            txt_Meta_1.Text = string.Empty;
            txt_Meta_2.Text = string.Empty;
            txt_Meta_3.Text = string.Empty;
            txt_Meta_4.Text = string.Empty;
            txt_Meta_5.Text = string.Empty;

            //limpiar variables globales
            id_global_meta_deshidratado = string.Empty;
            id_global_meta_empacado = string.Empty;
            id_global_meta_inspec = string.Empty;
            id_global_meta_evaporado = string.Empty;
            id_global_meta_revolturas = string.Empty;
            id_global_meta_polvos = string.Empty;

            //enfocar controles
            cmb_area.Focus();
            txt_op.Focus();
        }
        private void btn_meta_save_Click(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == -1)
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            /////Deshidratado
            if (cmb_area.SelectedIndex == 0 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_3.Text == string.Empty || txt_Meta_5.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 0)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_deshidratado))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_deshidratado);

                    queryInsertUpdate = "UPDATE public.\"Deshidratado\" SET \"OP\" = @OP, \"No_box_hr\" = @no_box_hr, \"Kg_fresco_hr\" = @kg_f_h, \"Relacion_fr_seco\" = @relacion_fs, \"Kg_seco_hr\" = @kg_s_h, \"Personal_idoneo\" = @personal_i WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@no_box_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@kg_f_h", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@relacion_fs", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@kg_s_h", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@personal_i", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_5.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {

                    // Verificar si ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Deshidratado\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Deshidratado\" (\"OP\", \"No_box_hr\", \"Kg_fresco_hr\", \"Relacion_fr_seco\", \"Kg_seco_hr\", \"Personal_idoneo\") VALUES (@OP, @No_box_h, @Kg_f_h, @Relacion_f_s, @kg_s_h, @Personal_i);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@No_box_h", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_f_h", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Relacion_f_s", System.Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@kg_s_h", System.Convert.ToDecimal(txt_Meta_4.Text)),
                    new NpgsqlParameter("@Personal_i", System.Convert.ToDecimal(txt_Meta_5.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_Deshitratado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_4.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_deshidratado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            /////Empacado
            ///
            if (cmb_area.SelectedIndex == 1 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 1)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_empacado))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_empacado);

                    queryInsertUpdate = "UPDATE public.\"Empacado\" SET \"OP\" = @OP, \"Personal_idoneo\" = @Personal_I, \"Kg_person_hr\" = @Kg_p_hr, \"Meta_kg_hr_line\" = @Meta_kg_hr_line WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Kg_p_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_line", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {

                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Empacado\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Empacado\" (\"OP\", \"Personal_idoneo\", \"Kg_person_hr\", \"Meta_kg_hr_line\") VALUES (@OP, @Personal_I, @Kg_person_hr, @Meta_kg_hr_line);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_person_hr", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_line", System.Convert.ToDecimal(txt_Meta_3.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_Empacado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_empacado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Inspeccion
            ///
            if (cmb_area.SelectedIndex == 2 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 2)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_inspec))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_inspec);

                    queryInsertUpdate = "UPDATE public.\"Inspeccion\" SET \"OP\" = @OP, \"Personal_idoneo\" = @Personal_I, \"Kg_person_hr\" = @Kg_p_hr, \"Meta_kg_hr_line\" = @Meta_kg_hr_line WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Kg_p_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_line", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Inspeccion\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Inspeccion\" (\"OP\", \"Personal_idoneo\", \"Kg_person_hr\", \"Meta_kg_hr_line\") VALUES (@OP, @Personal_I, @Kg_person_hr, @Meta_kg_hr_line);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Kg_person_hr", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_line", System.Convert.ToDecimal(txt_Meta_3.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_inspec();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_inspec = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Evaporado
            ///
            if (cmb_area.SelectedIndex == 3 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 3)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_evaporado))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_evaporado);

                    queryInsertUpdate = "UPDATE public.\"Evaporado\" SET \"OP\" = @OP, \"Personal_idoneo\" = @Personal_I, \"Capacidad_trompos\" = @Capacidad_T, \"Cantidad_trompos\" = @Cantidad_T, \"Meta_kg_hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Capacidad_T", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_T", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Evaporado\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Evaporado\" (\"OP\", \"Personal_idoneo\", \"Capacidad_trompos\", \"Cantidad_trompos\", \"Meta_kg_hr\") VALUES (@OP, @Personal_I, @Capacidad_trompos, @Cantidad_trompos, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_trompos", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_trompos", System.Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", System.Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                if (result > 0)
                {
                    actualiza_grid_evaporado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_evaporado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Grind
            ///
            if (cmb_area.SelectedIndex == 4 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 4)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_grind))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_grind);

                    queryInsertUpdate = "UPDATE public.\"Grind\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Capacidad_Molino\" = @Capacidad_m, \"Cantidad_molinos\" = @Cantidad_m, \"Meta_Kg_hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Capacidad_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Grind\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Grind\" (\"OP\", \"Personal_Idoneo\", \"Capacidad_Molino\", \"Cantidad_molinos\", \"Meta_Kg_hr\") VALUES (@OP, @Personal_I, @Capacidad_m, @Cantidad_m, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_m", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_m", System.Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", System.Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_grid_grind();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_grind = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Revolturas
            ///
            if (cmb_area.SelectedIndex == 5 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 5)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_revolturas))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_revolturas);

                    queryInsertUpdate = "UPDATE public.\"Revolturas\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Cap_Trompo_machin\" = @Cap_Trompo_m, \"Canti_Trompo_machin\" = @Cantidad_Trompo_m, \"Meta_Kg_Hr\" = @Meta_kg_hr WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Cap_Trompo_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_Trompo_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Revolturas\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Revolturas\" (\"OP\", \"Personal_Idoneo\", \"Cap_Trompo_machin\", \"Canti_Trompo_machin\", \"Meta_Kg_Hr\") VALUES (@OP, @Personal_I, @Capacidad_Trompo_m, @Cantidad_Trompo_m, @Meta_kg_hr);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_Trompo_m", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_Trompo_m", System.Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", System.Convert.ToDecimal(txt_Meta_4.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_revolturas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_revolturas = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Polvos
            ///
            if (cmb_area.SelectedIndex == 6 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 6)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_polvos))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_polvos);

                    queryInsertUpdate = "UPDATE public.\"Polvos\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Meta_kg_hr_hum\" = @Meta_kg_hr_hum, \"Meta_kg_hr_idon\" = @Meta_kg_hr_idon WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_hum", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr_idon", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Polvos\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Polvos\" (\"OP\", \"Personal_Idoneo\", \"Meta_kg_hr_hum\", \"Meta_kg_hr_idon\") VALUES (@OP, @Personal_I, @Meta_kg_hr_hum, @Meta_kg_hr_idon);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_hum", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Meta_kg_hr_idon", System.Convert.ToDecimal(txt_Meta_3.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                if (result > 0)
                {
                    actualiza_polvos();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_polvos = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
            ///Máquinas
            ///
            if (cmb_area.SelectedIndex == 7 && (txt_op.Text == string.Empty || txt_Meta_1.Text == string.Empty || txt_Meta_2.Text == string.Empty || txt_Meta_3.Text == string.Empty || txt_Meta_5.Text == string.Empty))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (cmb_area.SelectedIndex == 7)
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_meta_maquinas))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idmeta = System.Convert.ToInt32(id_global_meta_maquinas);

                    queryInsertUpdate = "UPDATE public.\"Maquinas\" SET \"OP\" = @OP, \"Personal_Idoneo\" = @Personal_I, \"Cap_Trompo_machin\" = @Cap_Trompo_m, \"Canti_Trompo_machin\" = @Cantidad_Trompo_m, \"Meta_Kg_Hr\" = @Meta_kg_hr, \"Kg_pz\" = @Kg_pz WHERE \"ID_OP\" = @ID_OP;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op.Text.ToUpper().Trim()
                        },
                        new NpgsqlParameter("@Personal_I", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_1.Text)
                        },
                        new NpgsqlParameter("@Cap_Trompo_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_2.Text)
                        },
                        new NpgsqlParameter("@Cantidad_Trompo_m", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_3.Text)
                        },
                        new NpgsqlParameter("@Meta_kg_hr", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_4.Text)
                        },
                        new NpgsqlParameter("@Kg_pz", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_Meta_5.Text)
                        },
                        new NpgsqlParameter("@ID_OP", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idmeta  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si el usuario ya existe
                    string queryCheckmeta = "SELECT COUNT(*) FROM public.\"Maquinas\" WHERE \"OP\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheck = dbHelper.ExecuteSelectQuery(queryCheckmeta, parametersCheck);
                    if (dtCheck != null && dtCheck.Rows.Count > 0 && System.Convert.ToInt32(dtCheck.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Maquinas\" (\"OP\", \"Personal_Idoneo\", \"Cap_Trompo_machin\", \"Canti_Trompo_machin\", \"Meta_Kg_Hr\", \"Kg_pz\") VALUES (@OP, @Personal_I, @Capacidad_Trompo_m, @Cantidad_Trompo_m, @Meta_kg_hr, @Kg_pz);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_op.Text.ToUpper().Trim()),
                    new NpgsqlParameter("@Personal_I", System.Convert.ToDecimal(txt_Meta_1.Text)),
                    new NpgsqlParameter("@Capacidad_Trompo_m", System.Convert.ToDecimal(txt_Meta_2.Text)),
                    new NpgsqlParameter("@Cantidad_Trompo_m", System.Convert.ToDecimal(txt_Meta_3.Text)),
                    new NpgsqlParameter("@Meta_kg_hr", System.Convert.ToDecimal(txt_Meta_4.Text)),
                    new NpgsqlParameter("@Kg_pz", System.Convert.ToDecimal(txt_Meta_5.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_maquinas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_maquinas = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
            }
        }
        private void dgv_metas_des_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 0;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_deshidratado = dgv_metas_des.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_des.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_des.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_2.Text = dgv_metas_des.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_3.Text = dgv_metas_des.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_4.Text = dgv_metas_des.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_5.Text = dgv_metas_des.Rows[e.RowIndex].Cells[6].Value.ToString();

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                
                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                txt_Meta_4.Enabled = false;
                txt_Meta_5.Enabled = false;

                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }
        private void txt_Meta_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }
        private void txt_Meta_1_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_2_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }
        private void txt_Meta_2_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_3_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }
        private void txt_Meta_3_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox.Focus();
                e.Cancel = true;
            }
        }

        private void txt_Meta_4_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }
        private void txt_Meta_4_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox.Focus();
                e.Cancel = true;
            }
        }
        private void txt_Meta_5_KeyPress(object sender, KeyPressEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }

                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }
        private void txt_Meta_5_Validating(object sender, CancelEventArgs e)
        {
            MaterialTextBox textBox = (MaterialTextBox)sender;
            string text = textBox.Text;

            if (!string.IsNullOrEmpty(text) && !System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox.Focus();
                e.Cancel = true;
            }
        }
        private void btn_meta_edit_Click(object sender, EventArgs e)
        {
            btn_meta_edit.Enabled = false;
            if (cmb_area.SelectedIndex == 0)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_3.Enabled = true;
                txt_Meta_5.Enabled = true;
                txt_op.Focus();
            }
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_op.Focus();
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5 || cmb_area.SelectedIndex == 6 || cmb_area.SelectedIndex == 7)
            {
                btn_meta_save.Enabled = true;
                btn_meta_cancel.Enabled = true;
                btn_meta_delete.Enabled = false;
                txt_op.Enabled = true;
                txt_Meta_1.Enabled = true;
                txt_Meta_2.Enabled = true;
                txt_Meta_3.Enabled = true;
                if (cmb_area.SelectedIndex == 7) { txt_Meta_5.Enabled = true; }
                txt_op.Focus();
            }
        }
        private void btn_meta_cancel_Click(object sender, EventArgs e)
        {
            btn_meta_save.Enabled = false;
            btn_meta_cancel.Enabled = false;
            btn_meta_delete.Enabled = false;
            limpiarCampos_meta();
            cmb_area.Enabled = false;
            txt_op.Enabled = false;
            txt_Meta_1.Enabled = false;
            txt_Meta_2.Enabled = false;
            txt_Meta_3.Enabled = false;
            txt_Meta_4.Enabled = false;
            txt_Meta_5.Enabled = false;
            btn_meta_edit.Enabled = false;
            id_global_meta_deshidratado = string.Empty;
            id_global_meta_empacado = string.Empty;
            id_global_meta_inspec = string.Empty;
        }
        private void btn_delete_user_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este usuario?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_user = System.Convert.ToInt32(id_global_users);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Usuarios""
                       WHERE ""ID_User"" = @idUser";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_user
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_users(); // Actualizar el DataGridView de usuarios
                    limpiarCampos(); // Limpiar los campos de texto
                    btn_save.Enabled = false;
                    btn_cancel.Enabled = false;
                    btn_delete_user.Enabled = false;
                    txt_no_emp.Enabled = false;
                    txt_usuario.Enabled = false;
                    txt_contra.Enabled = false;
                    btn_edit.Enabled = false;
                    id_global_users = string.Empty;

                    cmb_nivel_user.Enabled = true;
                    cmb_nivel_user.Focus(); // Enfocar el ComboBox de nivel de usuario
                    cmb_nivel_user.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el usuario", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }
        private void btn_meta_delete_Click(object sender, EventArgs e)
        {
            ///deshidratado
            ///
            if (cmb_area.SelectedIndex == 0 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Deshidratado?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_deshidratado);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Deshidratado""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };
                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_Deshitratado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_4.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_deshidratado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///empacado
            ///
            if (cmb_area.SelectedIndex == 1 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Empacado?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_empacado);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Empacado""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_Empacado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_empacado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///inspeccion
            ///
            if (cmb_area.SelectedIndex == 2 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Inspección?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_inspec);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Inspeccion""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_inspec();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_inspec = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Evaporado
            ///
            if (cmb_area.SelectedIndex == 3 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Evaporado?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_evaporado);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Evaporado""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_evaporado();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_evaporado = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Grind
            ///
            if (cmb_area.SelectedIndex == 4 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Grind?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_grind);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Grind""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_grid_grind();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_grind = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Revolturas
            ///
            if (cmb_area.SelectedIndex == 5 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Revolturas?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_revolturas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Revolturas""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_revolturas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_revolturas = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Polvos
            ///
            if (cmb_area.SelectedIndex == 5 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Revolturas?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_revolturas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Polvos""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_polvos();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_polvos = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Máquinas
            ///
            if (cmb_area.SelectedIndex == 7 && (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar este OP de la tabla Máquinas?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                int id_OP = System.Convert.ToInt32(id_global_meta_maquinas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Maquinas""
                       WHERE ""ID_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };

                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_maquinas();
                    limpiarCampos_meta();
                    btn_meta_save.Enabled = false;
                    btn_meta_cancel.Enabled = false;
                    btn_meta_delete.Enabled = false;
                    txt_op.Enabled = false;
                    txt_Meta_1.Enabled = false;
                    txt_Meta_2.Enabled = false;
                    txt_Meta_3.Enabled = false;
                    txt_Meta_5.Enabled = false;
                    btn_meta_edit.Enabled = false;
                    id_global_meta_maquinas = string.Empty;

                    cmb_area.Enabled = true;
                    cmb_area.Focus();
                    cmb_area.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar el OP", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void txt_Meta_1_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 0)
            {
                
                if (decimal.TryParse(txt_Meta_1.Text, out decimal valor))
                {
                    // Fórmula: (valor * 380) - (valor * 380 * 10%)
                    decimal resultado = (valor * 380) - (valor * 380 * 0.10m);

                    txt_Meta_2.Text = resultado.ToString("0.00"); // 2 decimales
                }
                else
                {
                    txt_Meta_2.Text = string.Empty; // Limpia si no es válido
                }
                kg_seco_hr();
            }
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                meta_hr_linea();
            }
        }
        private void kg_seco_hr()
        {
            if (decimal.TryParse(txt_Meta_3.Text, out decimal valor2) && decimal.TryParse(txt_Meta_2.Text, out decimal valor1))
            {
                // Fórmula: KG FRESCO POR HORA/RELACION FRESCO-SECO
                decimal resultado = valor1 / valor2;

                txt_Meta_4.Text = resultado.ToString("0.00"); // 2 decimales
            }
            else
            {
                txt_Meta_4.Text = string.Empty; // Limpia si no es válido
            }
        }

        private void txt_Meta_3_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 0)
            {
                kg_seco_hr();
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5 || cmb_area.SelectedIndex == 7)
            {
                meta_hr();
            }
        }
        private void meta_hr()
        {
            if (decimal.TryParse(txt_Meta_3.Text, out decimal valor2) && decimal.TryParse(txt_Meta_2.Text, out decimal valor1))
            {
                // Fórmula: KG FRESCO POR HORA/RELACION FRESCO-SECO
                decimal resultado = valor1 * valor2;

                txt_Meta_4.Text = resultado.ToString("0.00"); // 2 decimales
            }
            else
            {
                txt_Meta_4.Text = string.Empty; // Limpia si no es válido
            }
        }
        private void txt_Meta_2_TextChanged(object sender, EventArgs e)
        {
            if (cmb_area.SelectedIndex == 1 || cmb_area.SelectedIndex == 2)
            {
                meta_hr_linea();
            }
            if (cmb_area.SelectedIndex == 3 || cmb_area.SelectedIndex == 4 || cmb_area.SelectedIndex == 5 || cmb_area.SelectedIndex == 7)
            {
                meta_hr();
            }
        }

        private void meta_hr_linea()
        {
            if (decimal.TryParse(txt_Meta_2.Text, out decimal valor2) && decimal.TryParse(txt_Meta_1.Text, out decimal valor1))
            {
                // Fórmula: KG FRESCO POR HORA/RELACION FRESCO-SECO
                decimal resultado = valor1 * valor2;

                txt_Meta_3.Text = resultado.ToString("0.00"); // 2 decimales
            }
            else
            {
                txt_Meta_3.Text = string.Empty; // Limpia si no es válido
            }
        }

        private void dgv_metas_emp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 1;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_empacado = dgv_metas_emp.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_2.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_emp.Rows[e.RowIndex].Cells[3].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_insp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 2;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_inspec = dgv_metas_insp.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_2.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_insp.Rows[e.RowIndex].Cells[3].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                
                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_Eva_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 3;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_evaporado = dgv_metas_Eva.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_Eva.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                
                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_Grind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 4;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_grind = dgv_metas_Grind.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_Grind.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                
                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void cmb_area_search_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_area_search.SelectedIndex != -1)
            {
                txt_OP_Search.Enabled = true;
                if (cmb_area_search.SelectedIndex == 0)
                {
                    tap_control_metas.SelectedIndex = 0;
                }
                if (cmb_area_search.SelectedIndex == 1)
                {
                    tap_control_metas.SelectedIndex = 1;
                }
                if (cmb_area_search.SelectedIndex == 2)
                {
                    tap_control_metas.SelectedIndex = 2;
                }
                if (cmb_area_search.SelectedIndex == 3)
                {
                    tap_control_metas.SelectedIndex = 3;
                }
                if (cmb_area_search.SelectedIndex == 4)
                {
                    tap_control_metas.SelectedIndex = 4;
                }
                if (cmb_area_search.SelectedIndex == 5)
                {
                    tap_control_metas.SelectedIndex = 5;
                }
            }
        }

        private void materialButton7_Click(object sender, EventArgs e)
        {
            //Deshidratado
            ///
            if (cmb_area_search.SelectedIndex == 0)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 1;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_des.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();
                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_des.ClearSelection();
                dgv_metas_des.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_des.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_des.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;

                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Empacado
            ///
            if (cmb_area_search.SelectedIndex == 1)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 2;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_emp.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_emp.ClearSelection();
                dgv_metas_emp.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_emp.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_emp.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;

                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Inspeccion
            ///
            if (cmb_area_search.SelectedIndex == 2)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 3;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_insp.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_insp.ClearSelection();
                dgv_metas_insp.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_insp.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_insp.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Evaporado
            ///
            if (cmb_area_search.SelectedIndex == 3)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 4;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_Eva.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_Eva.ClearSelection();
                dgv_metas_Eva.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_Eva.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_Eva.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Grind
            ///
            if (cmb_area_search.SelectedIndex == 4)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 5;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_Grind.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno                          
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_Grind.ClearSelection();
                dgv_metas_Grind.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_Grind.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_Grind.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            ///Revolturas
            ///
            if (cmb_area_search.SelectedIndex == 5)
            {
                //dgv activo
                filtroUsuariosActivo_OP_dgv_activo = 6;
                // Si el filtro está activo, limpiar filtro y campos de búsqueda
                if (filtroUsuariosActivo_OP)
                {
                    // Mostrar todas las filas del DataGridView
                    foreach (DataGridViewRow row in dgv_metas_revolturas.Rows)
                    {
                        row.Visible = true;
                    }

                    // Restablecer el estado del filtro
                    filtroUsuariosActivo_OP = false;
                }

                string OP = txt_OP_Search.Text.Trim();

                //validar que el campo área este lleno
                if (cmb_area_search.SelectedIndex == -1)
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo Área para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //validar que el campo OP este lleno
                if (string.IsNullOrEmpty(OP))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Deseleccionar cualquier celda y fila antes de filtrar
                dgv_metas_revolturas.ClearSelection();
                dgv_metas_revolturas.CurrentCell = null;

                // Mover el CurrencyManager a una posición válida
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_metas_revolturas.DataSource];
                if (cm != null && cm.Count > 0)
                    cm.Position = -1;

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_metas_revolturas.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                filtroUsuariosActivo = hayFiltro;
                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron resultados con los criterios de búsqueda.",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void materialButton5_Click(object sender, EventArgs e)
        {
            limpiar_filtros_OP();
        }

        private void calcular_meta_programada()
        {
            double hr_programada, meta_programada, meta_x_hr, kg_pz_double;

            hr_programada = double.TryParse(Txt_Read_1.Text, out hr_programada) ? hr_programada : 0;
            meta_x_hr = double.TryParse(Txt_meta.Text, out meta_x_hr) ? meta_x_hr : 0;
            if (cb_Area.SelectedIndex == 8)
            {
                kg_pz_double = double.TryParse(Kg_pz_str, out kg_pz_double) ? kg_pz_double : 0;
                meta_programada = (meta_x_hr* kg_pz_double) *hr_programada;
            }
            else
            {
                meta_programada = hr_programada * meta_x_hr;
            }
            
            Txt_Read_2.Text = meta_programada.ToString("0.##");
        }
        private void calcular_turno()
        {
            if (cb_Area.SelectedIndex != -1)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(Mask_txt_hr1.Text) && !string.IsNullOrWhiteSpace(Mask_txt_hr2.Text)
                        && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :")
                    {
                        // Crear copias de las horas para no modificar las originales
                        DateTime inicio = horaInicio;
                        DateTime fin = horaFin;

                        // CORRECCIÓN: Determinar si la hora fin es del día siguiente
                        // Si fin es menor que inicio, asumimos que es del día siguiente
                        // Esto funciona para cualquier caso de horas que crucen la medianoche
                        if (fin <= inicio)
                        {
                            fin = fin.AddDays(1);
                        }

                        // Calcular diferencia inicial
                        TimeSpan diferencia = fin - inicio;

                        // Restar minutos energía
                        if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text) && int.TryParse(txt_Tiempo_energia.Text, out int minutosEnergia))
                        {
                            diferencia = diferencia.Subtract(TimeSpan.FromMinutes(minutosEnergia));
                        }

                        //validar si el turno es menor a 6 horas
                        // Convertir la diferencia a minutos totales (como double)
                        double minutosTotales = diferencia.TotalMinutes;

                        // Comparar con 360 minutos (6 horas = 360 minutos)
                        if (minutosTotales < 360)
                        {
                            txt_Tiempo_comida.Text = "0";
                        }

                        // Restar minutos comida
                        if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && int.TryParse(txt_Tiempo_comida.Text, out int minutosComida1))
                        {
                            diferencia = diferencia.Subtract(TimeSpan.FromMinutes(minutosComida1));

                            // Asegurar que no sea negativo después de restar comida
                            if (diferencia.TotalMinutes < 0)
                            {
                                diferencia = TimeSpan.Zero;
                            }
                        }

                        // Mostrar horas totales (con decimales si hay minutos)
                        Txt_Read_1.Text = diferencia.TotalHours.ToString("0.##");

                        // Calcular la nueva diferencia restando los minutos
                        TimeSpan diferenciaConDescuento = diferencia;

                        // Restar minutos mecánicos
                        if (!string.IsNullOrEmpty(txt_TM_mecanico.Text) && int.TryParse(txt_TM_mecanico.Text, out int minutosMecanico))
                        {
                            diferenciaConDescuento = diferenciaConDescuento.Subtract(TimeSpan.FromMinutes(minutosMecanico));
                        }

                        // Restar minutos operativos
                        if (!string.IsNullOrEmpty(txt_TM_operativo.Text) && int.TryParse(txt_TM_operativo.Text, out int minutosOperativo))
                        {
                            diferenciaConDescuento = diferenciaConDescuento.Subtract(TimeSpan.FromMinutes(minutosOperativo));
                        }

                        // Asegurar que no sea negativo
                        if (diferenciaConDescuento.TotalMinutes < 0)
                        {
                            diferenciaConDescuento = TimeSpan.Zero;
                        }

                        if (!string.IsNullOrEmpty(Txt_meta.Text) && cb_Area.SelectedIndex != 4)
                        {
                            calcular_meta_programada();
                        }

                        // Mostrar el resultado con descuento en Txt_Read_3
                        Txt_Read_3.Text = diferenciaConDescuento.TotalHours.ToString("0.##");
                    }
                    btn_save_ficha.Enabled = true;
                }
                catch (Exception ex)
                {
                    if (ex is FormatException)
                    {
                        btn_save_ficha.Enabled = false;
                    }
                }
            }
        }
        //private void calcular_turno()
        //{
        //    if (cb_Area.SelectedIndex != -1)
        //    {
        //        try
        //        {
        //            if (!string.IsNullOrWhiteSpace(Mask_txt_hr1.Text) && !string.IsNullOrWhiteSpace(Mask_txt_hr2.Text)
        //                && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :")
        //            {
        //                // Crear copias de las horas para no modificar las originales
        //                DateTime inicio = horaInicio;
        //                DateTime fin = horaFin;

        //                // Verificar si la hora de fin es menor que la de inicio
        //                // Pero considerar que si la diferencia es muy grande (>12 horas),
        //                // probablemente es porque pasó al día siguiente
        //                if (fin < inicio)
        //                {
        //                    // Si la diferencia es más de 12 horas en negativo, asumimos que es del día siguiente
        //                    if ((inicio - fin).TotalHours > 12)
        //                    {
        //                        fin = fin.AddDays(1);
        //                    }
        //                }

        //                // Calcular diferencia inicial
        //                TimeSpan diferencia = fin - inicio;

        //                // Restar minutos energia
        //                if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text) && int.TryParse(txt_Tiempo_energia.Text, out int minutosEnergia))
        //                {
        //                    diferencia = diferencia.Subtract(TimeSpan.FromMinutes(minutosEnergia));
        //                }
        //                //validar si el turno es menor a 6 horas
        //                // Convertir la diferencia a minutos totales (como double)
        //                double minutosTotales = diferencia.TotalMinutes;

        //                // Comparar con 360 minutos (6 horas = 360 minutos)
        //                if (minutosTotales < 360)
        //                {
        //                    txt_Tiempo_comida.Text = "0";
        //                }
        //                // Restar minutos comida
        //                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && int.TryParse(txt_Tiempo_comida.Text, out int minutosComida1))
        //                {
        //                    diferencia = diferencia.Subtract(TimeSpan.FromMinutes(minutosComida1));
        //                }

        //                // Mostrar horas totales (con decimales si hay minutos)
        //                Txt_Read_1.Text = diferencia.TotalHours.ToString("0.##");

        //                // Calcular la nueva diferencia restando los minutos
        //                TimeSpan diferenciaConDescuento = diferencia;

        //                // Restar minutos mecánicos
        //                if (!string.IsNullOrEmpty(txt_TM_mecanico.Text) && int.TryParse(txt_TM_mecanico.Text, out int minutosMecanico))
        //                {
        //                    diferenciaConDescuento = diferenciaConDescuento.Subtract(TimeSpan.FromMinutes(minutosMecanico));
        //                }

        //                // Restar minutos operativos
        //                if (!string.IsNullOrEmpty(txt_TM_operativo.Text) && int.TryParse(txt_TM_operativo.Text, out int minutosOperativo))
        //                {
        //                    diferenciaConDescuento = diferenciaConDescuento.Subtract(TimeSpan.FromMinutes(minutosOperativo));
        //                }

        //                // Asegurar que no sea negativo
        //                if (diferenciaConDescuento.TotalMinutes < 0)
        //                {
        //                    diferenciaConDescuento = TimeSpan.Zero;
        //                }

        //                if (!string.IsNullOrEmpty(Txt_meta.Text) && cb_Area.SelectedIndex != 4)
        //                {
        //                    calcular_meta_programada();
        //                }

        //                // Mostrar el resultado con descuento en Txt_Read_3
        //                Txt_Read_3.Text = diferenciaConDescuento.TotalHours.ToString("0.##");
        //            }
        //            btn_save_ficha.Enabled = true;
        //        }
        //        catch (Exception ex)
        //        {
        //            if (ex is FormatException)
        //            {
        //                btn_save_ficha.Enabled = false;
        //            }
        //        }
        //    }
        //}
        private void cb_Turno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_Turno.SelectedIndex != -1)
            {
                if (!editar) { reiniciarCampos_comboturno(); }
                if (cb_Turno.SelectedIndex == 0)
                {
                    Mask_txt_hr1.Text = "07:00";
                    Mask_txt_hr2.Text = "15:00";

                    calcular_turno();
                }
                if (cb_Turno.SelectedIndex == 1)
                {
                    Mask_txt_hr1.Text = "15:00";
                    Mask_txt_hr2.Text = "22:30";

                    calcular_turno();
                }
                if (cb_Turno.SelectedIndex == 2)
                {
                    Mask_txt_hr1.Text = "22:30";
                    Mask_txt_hr2.Text = "07:00";

                    calcular_turno();
                }
                Mask_txt_hr1.Enabled = true;
                Mask_txt_hr2.Enabled = true;
                txt_Tiempo_energia.Enabled = true;
                txt_Tiempo_comida.Enabled = true;
                dgv_mecanico.Enabled = true;
                dgv_operativo.Enabled = true;
                if (cb_Area.SelectedIndex == 0)
                {
                    if (!editar) 
                    { 
                        Txt_1.Enabled = true; 
                    } 
                    else 
                    { 
                        Txt_1.Enabled = false; 
                    }
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    Txt_6.Enabled = true;
                    Txt_7.Enabled = true;
                    Txt_8.Enabled = true;
                    Txt_9.Enabled = true;
                    Txt_10.Enabled = true;
                    Txt_11.Enabled = true;

                }
                if (cb_Area.SelectedIndex == 1)
                {
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    Txt_6.Enabled = true;
                }
                if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7)
                {
                    Txt_1.Enabled = true;
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    if (cb_Area.SelectedIndex == 4)
                    {
                        cb_proceso.Enabled = true;
                    }
                    if (cb_Area.SelectedIndex == 6)
                    {
                        Txt_6.Enabled = true;
                        Txt_7.Enabled = true;
                    }
                }
                if (cb_Area.SelectedIndex == 8)
                {
                    Txt_1.Enabled = true;
                    Txt_2.Enabled = true;
                    Txt_3.Enabled = true;
                    Txt_4.Enabled = true;
                    Txt_5.Enabled = true;
                    Txt_6.Enabled = true;
                    Txt_7.Enabled = true;
                    Txt_8.Enabled = true;
                }
            }
        }
        private bool ValidarTodasLasFilas_operativo()
        {
            int count_vacios = 0;
            foreach (DataGridViewRow row in dgv_operativo.Rows)
            {
                if (row.IsNewRow) continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    {
                        count_vacios++;
                    }
                }
            }

            if (count_vacios > 0)
            {
                return false;
            }
            else 
            {
                return true;
            }
        }
        private bool ValidarTodasLasFilas_mecanico()
        {
            int count_vacios = 0;
            foreach (DataGridViewRow row in dgv_mecanico.Rows)
            {
                if (row.IsNewRow) continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()))
                    {
                        count_vacios++;
                    }
                }
            }

            if (count_vacios > 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void btn_save_ficha_Click(object sender, EventArgs e)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            string Variable_nombre = lbl_nom2.Text;
            string variable_num = lbl_no_emp2.Text;
            string jefe = cb_jefe_turno.Text;
            if (editar) 
            {
                // Consulta para buscar donde OP = valor_buscado
                string query2 = "SELECT \"ID_User\" FROM public.\"Usuarios\" where \"No_Empleado\" = @valorBuscado";

                // Crear parámetro
                NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = variable_num
                        }
                };
                // Ejecutar consulta
                System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                string id_supervisor = string.Empty;
                // Verificar si se encontraron resultados
                if (dt2 != null && dt2.Rows.Count > 0)
                {
                    id_supervisor = dt2.Rows[0]["ID_User"].ToString();
                }
                id_supervisor_global = System.Convert.ToInt32(id_supervisor);
            }

            int idUserSeleccionado = System.Convert.ToInt32(cb_jefe_turno.SelectedValue);
            string cuerpoHtml1 = $@"
                <html>
                <body style='font-family: Arial; font-size: 14px; color: #333;'>

                    <div style='border: 1px solid #ddd; padding: 20px; border-radius: 8px; background: #f7f7f7;'>

                        <h2 style='color: #2c3e50;'>Reporte del Sistema Tablero</h2>

                        <p>
                            Estimado equipo,<br><br>
                            El sistema <strong>Tablero</strong> informa que el supervisor 
                            <strong>{Variable_nombre}</strong>, con número de empleado 
                            <strong>{variable_num}</strong> ha registrado la información correspondiente 
                            al jefe de turno <strong>{jefe}</strong>, correspondiente al turno ";

            ////TUNEL/ SUMERGIDOR
            if (cb_Area.SelectedIndex == 0)
            {
                try
                {
                    int idUsuarioActual = id_user;
                   

                    if (cb_OP.SelectedIndex != -1 && !string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) 
                        && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) 
                        && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) 
                        && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(Txt_6.Text) && !string.IsNullOrEmpty(Txt_7.Text) 
                        && !string.IsNullOrEmpty(Txt_8.Text) && !string.IsNullOrEmpty(Txt_9.Text) && !string.IsNullOrEmpty(Txt_10.Text)
                        && !string.IsNullOrEmpty(cb_jefe_turno.Text)&& ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                    {
                        // Obtener datos de los TextBox
                        DateTime fecha = dtp1.Value; // Tu MetroDateTime
                        int turno = System.Convert.ToInt32(cb_Turno.Text);
                        string lote = Txt_1.Text;
                        string op = cb_OP.Text;
                        decimal kgEnterProceso = System.Convert.ToDecimal(Txt_2.Text);
                        decimal kgFrescosEnterSe = System.Convert.ToDecimal(Txt_Read_4.Text);
                        decimal mermaCanica = System.Convert.ToDecimal(Txt_3.Text);
                        decimal mermaPodrido = System.Convert.ToDecimal(Txt_4.Text);
                        decimal mermaTina = System.Convert.ToDecimal(Txt_5.Text);
                        decimal mermaPiso = System.Convert.ToDecimal(Txt_6.Text);
                        decimal mermaCanaletas = System.Convert.ToDecimal(Txt_7.Text);
                        decimal mermaLavadoBandas = System.Convert.ToDecimal(Txt_8.Text);
                        decimal cascaraCarrete = System.Convert.ToDecimal(Txt_10.Text);
                        int personal_Op = System.Convert.ToInt32(Txt_9.Text);
                        decimal hr_pro = System.Convert.ToDecimal(Txt_Read_1.Text);
                        decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                        decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                        string area = cb_Area.Text;
                        decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                        

                        // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                        TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                        TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                        string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                        if (editar)
                        {
                            idUsuarioActual = id_supervisor_global;
                            updateFicha(dbHelper, idUsuarioActual, fecha, turno, null, op,
                            kgEnterProceso, kgFrescosEnterSe, mermaCanica, mermaPodrido, mermaTina, mermaPiso,
                            mermaCanaletas, mermaLavadoBandas, cascaraCarrete, hrInicio, hrFin, personal_Op, hr_pro, hr_efec, meta_kg, null, meta, 0, idUserSeleccionado);
                            //enviar correo
                            //Crear lista de valores
                            List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Lote", lote),
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg Entrada (Proceso)", kgEnterProceso.ToString()),
                                new KeyValuePair<string, string>("Kg Frescos de Entrada a secador", kgFrescosEnterSe.ToString()),
                                new KeyValuePair<string, string>("Kg Canica", mermaCanica.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Podrido", mermaPodrido.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Tina", mermaTina.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Piso", mermaPiso.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Canaletas", mermaCanaletas.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Lavado de Bandas", mermaLavadoBandas.ToString()),
                                new KeyValuePair<string, string>("Kg Cáscara de Carrete", cascaraCarrete.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", personal_Op.ToString()),
                                new KeyValuePair<string, string>("Horas Programadas", hr_pro.ToString()),
                                new KeyValuePair<string, string>("Kg de Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString())
                            };
                            string tablaGenerada = GenerarTablaValores(valores);
                            string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                            string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                            string resultado = GenerarTablaDetallesOP();
                            string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>
                                    <br>
                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                            cuerpoHtml1 += cuerpoHtml2;
                            conf_email(cuerpoHtml1);
                            //fin correo
                        }
                        else
                        {
                            // Verificar si ya existe
                            string queryChecklote = "SELECT COUNT(*) FROM public.\"Ficha\" WHERE \"Lote\"  ILIKE @Lote and \"OP\" = @op;";
                            NpgsqlParameter[] parametersLote = new NpgsqlParameter[]
                            {
                                new NpgsqlParameter("@Lote", lote),
                                new NpgsqlParameter("@op", op)
                            };
                            System.Data.DataTable dtLote = dbHelper.ExecuteSelectQuery(queryChecklote, parametersLote);
                            if (dtLote != null && dtLote.Rows.Count > 0 && System.Convert.ToInt32(dtLote.Rows[0][0]) > 0)
                            {
                                MetroFramework.MetroMessageBox.Show(this, "El Lote ya existe con el OP selecionado. Por favor, elija otro lote.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            // Insertar en tabla Ficha y obtener el ID_Ficha generado
                            int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, lote, op,
                                kgEnterProceso, kgFrescosEnterSe, mermaCanica, mermaPodrido, mermaTina, mermaPiso,
                                mermaCanaletas, mermaLavadoBandas, cascaraCarrete, hrInicio, hrFin, personal_Op, hr_pro, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);

                            if (idFicha > 0)
                            {

                                // Insertar en tablas relacionadas
                                InsertarTiemposMuertos(dbHelper, idFicha);

                                //enviar correo
                                //Crear lista de valores
                                List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                                {
                                    new KeyValuePair<string, string>("Lote", lote),
                                    new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                    new KeyValuePair<string, string>("Kg Entrada (Proceso)", kgEnterProceso.ToString()),
                                    new KeyValuePair<string, string>("Kg Frescos de Entrada a secador", kgFrescosEnterSe.ToString()),
                                    new KeyValuePair<string, string>("Kg Canica", mermaCanica.ToString()),
                                    new KeyValuePair<string, string>("Kg Merma Podrido", mermaPodrido.ToString()),
                                    new KeyValuePair<string, string>("Kg Merma Tina", mermaTina.ToString()),
                                    new KeyValuePair<string, string>("Kg Merma Piso", mermaPiso.ToString()),
                                    new KeyValuePair<string, string>("Kg Merma Canaletas", mermaCanaletas.ToString()),
                                    new KeyValuePair<string, string>("Kg Merma Lavado de Bandas", mermaLavadoBandas.ToString()),
                                    new KeyValuePair<string, string>("Kg Cáscara de Carrete", cascaraCarrete.ToString()),
                                    new KeyValuePair<string, string>("Personal Operativo", personal_Op.ToString()),
                                    new KeyValuePair<string, string>("Horas Programadas", hr_pro.ToString()),
                                    new KeyValuePair<string, string>("Kg de Meta Programada", meta_kg.ToString()),
                                    new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString())
                                };
                                string valorBuscado = cb_OP.Text;
                                // Consulta para buscar donde OP = valor_buscado
                                string query = "SELECT \"Orden_Produccion\", \"Producto\", \"Medida\", \"Descripcion\", " +
                                    "\"Especificacion\", \"Ingredientes\", \"Humedad\", \"Comercio\", \"Manzana\", \"Analisis\", " +
                                    "\"Area_Proceso\", \"OP_Origen\", \"Destino\", \"Clasificacion\" \r\nFROM public.\"Detalles_OP\" " +
                                    "where \"Orden_Produccion\" = @valorBuscado";

                                // Crear parámetro
                                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                                {
                                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                                };

                                // Ejecutar consulta
                                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                                string tablaGenerada = GenerarTablaValores(valores);
                                // Generar tabla de tiempos muertos mecánicos
                                string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                                string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                                string resultado = GenerarTablaDetallesOP();
                                string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>
                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                                cuerpoHtml1 += cuerpoHtml2;
                                conf_email(cuerpoHtml1);
                                //fin correo
                                MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                    "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                cb_Area.SelectedIndex = -1;
                                reiniciarCampos();
                                cb_Area.Focus();
                            }
                            else
                            {
                                MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                    "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    else
                    {
                        MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                            "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, $"Error al guardar: {ex.Message}",
                                                            "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (cb_Area.SelectedIndex == 1)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 
                    && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) 
                    && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(Txt_6.Text) 
                    && !string.IsNullOrEmpty(cb_jefe_turno.Text) && ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                {
                    //Despegue
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = System.Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal KgProdSeco = System.Convert.ToDecimal(Txt_2.Text);
                    decimal MermaKgSeco = System.Convert.ToDecimal(Txt_3.Text);
                    decimal KgFueraSpec = System.Convert.ToDecimal(Txt_4.Text);
                    decimal KgResecar = System.Convert.ToDecimal(Txt_5.Text);
                    int PersonalOpe = System.Convert.ToInt32(Txt_6.Text);
                    decimal hr_programadas = System.Convert.ToDecimal(Txt_Read_1.Text);
                    decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                    decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                    string textoPorcentCumplimiento = Txt_Read_5.Text.Replace("%", "").Trim();
                    decimal PorcentCumplimiento = System.Convert.ToDecimal(textoPorcentCumplimiento) / 100m;
                    //decimal Kg_secos_meta = System.Convert.ToDecimal(Txt_Read_6.Text);
                    decimal Relacion_Fresco_seco = System.Convert.ToDecimal(Txt_Read_7.Text);//-------------------------------------adjuntar a resumen correo
                    string textoFTT = Txt_Read_8.Text.Replace("%", "").Trim();
                    decimal FTT = System.Convert.ToDecimal(textoFTT) / 100m;
                    string area = cb_Area.Text;
                    decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                    string lote = radMultiColumnComboBox1.Text;
                    decimal kgFrescosEnterSe = System.Convert.ToDecimal(Txt_Read_4.Text);//-------------------------------------adjuntar a resumen correo
                    string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);

                    if (editar)
                    {
                        idUsuarioActual = id_supervisor_global;
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                            KgFueraSpec, KgResecar, PorcentCumplimiento, 0, Relacion_Fresco_seco, FTT,
                            KgProdSeco, MermaKgSeco, kgFrescosEnterSe, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0, idUserSeleccionado);
                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Lote", lote),
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kilos Producto Seco", KgProdSeco.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Secos", MermaKgSeco.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraSpec.ToString()),
                                new KeyValuePair<string, string>("Kg para Resecar", KgResecar.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Kg Frescos Entrada a secador", kgFrescosEnterSe.ToString()),
                                new KeyValuePair<string, string>("Cumplimiento de Metas (%)", textoPorcentCumplimiento),
                                new KeyValuePair<string, string>("Relación Fresco-Seco", Relacion_Fresco_seco.ToString()),
                                new KeyValuePair<string, string>("FTT", textoFTT)
                            };
                        string Kgproducidosxhr = (KgProdSeco/ hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProdSeco.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr),
                                new KeyValuePair<string, string>("Kg que ingresan a Túnel", kgFrescosEnterSe.ToString()),
                                new KeyValuePair<string, string>("Relación Fresco/Seco", Relacion_Fresco_seco.ToString())
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                         <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, lote,
                            KgFueraSpec, KgResecar, PorcentCumplimiento, 0, Relacion_Fresco_seco, FTT,
                            KgProdSeco, MermaKgSeco, kgFrescosEnterSe, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);
                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Lote", lote),
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kilos Producto Seco", KgProdSeco.ToString()),
                                new KeyValuePair<string, string>("Kg Merma Secos", MermaKgSeco.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraSpec.ToString()),
                                new KeyValuePair<string, string>("Kg para Resecar", KgResecar.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Kg Frescos Entrada a secador", kgFrescosEnterSe.ToString()),
                                new KeyValuePair<string, string>("Cumplimiento de Metas (%)", textoPorcentCumplimiento),
                                new KeyValuePair<string, string>("Relación Fresco-Seco", Relacion_Fresco_seco.ToString()),
                                new KeyValuePair<string, string>("FTT", textoFTT)
                            };
                        string Kgproducidosxhr = (KgProdSeco / hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProdSeco.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr),
                                new KeyValuePair<string, string>("Kg que ingresan a Túnel", kgFrescosEnterSe.ToString()),
                                new KeyValuePair<string, string>("Relación Fresco/Seco", Relacion_Fresco_seco.ToString())
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                       <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo

                        if (idFicha > 0)
                        {
                            // Ahora hacer el UPDATE usando el parámetro @loteterminado
                            string updateQuery = @"UPDATE public.""Ficha"" 
                              SET ""Terminado_Tunel"" = true 
                              WHERE ""Lote"" = @loteterminado;";

                            NpgsqlParameter[] updateParameters = new NpgsqlParameter[]
                            {
                                new NpgsqlParameter("@loteterminado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = lote ?? (object)DBNull.Value }
                            };

                            int updateResult = dbHelper.ExecuteNonQuery(updateQuery, updateParameters);

                            // Insertar en tablas relacionadas
                            InsertarTiemposMuertos(dbHelper, idFicha);

                            MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            cb_Area.SelectedIndex = -1;
                            reiniciarCampos();
                            cb_Area.Focus();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 2)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 
                    && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) 
                    && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) 
                    && !string.IsNullOrEmpty(cb_jefe_turno.Text) && ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                {
                    //Evaporado
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = System.Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = System.Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = System.Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = System.Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = System.Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = System.Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = System.Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = System.Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                    string textoPorcent_Aumento_Hume = Txt_Read_5.Text.Replace("%", "").Trim();
                    string area = cb_Area.Text;
                    decimal Porcent_Aumento_Hume = System.Convert.ToDecimal(textoPorcent_Aumento_Hume) / 100m;
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                    string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                    if (editar)
                    {
                        idUsuarioActual = id_supervisor_global;
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        Porcent_Aumento_Hume, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0, idUserSeleccionado);

                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado),
                                new KeyValuePair<string, string>("Aumento por humedad (%)", textoPorcent_Aumento_Hume)
                            };
                        string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        Porcent_Aumento_Hume, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);

                        if (idFicha > 0)
                        {
                            // Insertar en tablas relacionadas
                            InsertarTiemposMuertos(dbHelper, idFicha);

                            //enviar correo
                            //Crear lista de valores
                            List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado),
                                new KeyValuePair<string, string>("Aumento por humedad (%)", textoPorcent_Aumento_Hume)
                            };
                            string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                            //lista de valores para resumen
                            List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                            string tablaResumen = GenerarTablaResumen(valoresResumen);
                            string tablaGenerada = GenerarTablaValores(valores);
                            // Generar tabla de tiempos muertos mecánicos
                            string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                            string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                            string resultado = GenerarTablaDetallesOP();
                            string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                      <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                            cuerpoHtml1 += cuerpoHtml2;
                            conf_email(cuerpoHtml1);
                            //fin correo

                            MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            cb_Area.SelectedIndex = -1;
                            reiniciarCampos();
                            cb_Area.Focus();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 7)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 
                    && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) 
                    && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) 
                    && !string.IsNullOrEmpty(cb_jefe_turno.Text) && ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                {
                    //Grind
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = System.Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = System.Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = System.Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = System.Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = System.Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = System.Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = System.Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = System.Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                    string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                    if (editar)
                    {
                        idUsuarioActual = id_supervisor_global;
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0, idUserSeleccionado);

                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado)
                            };
                        string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);
                        if (idFicha > 0)
                        {
                            // Insertar en tablas relacionadas
                            InsertarTiemposMuertos(dbHelper, idFicha);

                            //enviar correo
                            //Crear lista de valores
                            List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado)
                            };
                            string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                            //lista de valores para resumen
                            List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                            string tablaResumen = GenerarTablaResumen(valoresResumen);
                            string tablaGenerada = GenerarTablaValores(valores);
                            // Generar tabla de tiempos muertos mecánicos
                            string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                            string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                            string resultado = GenerarTablaDetallesOP();
                            string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                            cuerpoHtml1 += cuerpoHtml2;
                            conf_email(cuerpoHtml1);
                            //fin correo
                            MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            cb_Area.SelectedIndex = -1;
                            reiniciarCampos();
                            cb_Area.Focus();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 4)
            {
                if (cb_proceso.SelectedIndex != -1 && !string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) 
                    && cb_Turno.SelectedIndex != -1 && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) 
                    && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) 
                    && !string.IsNullOrEmpty(Txt_5.Text) && !string.IsNullOrEmpty(cb_jefe_turno.Text) && ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                {
                    //Inspeccion
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = System.Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = System.Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = System.Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = System.Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = System.Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = System.Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = System.Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = System.Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                    string proceso = cb_proceso.Text;
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                    string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                    if (editar)
                    {
                        idUsuarioActual = id_supervisor_global;
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, proceso,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, null, meta, 0, idUserSeleccionado);

                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado)
                            };
                        string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, proceso,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        0, 0, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);
                        if (idFicha > 0)
                        {
                            // Insertar en tablas relacionadas
                            InsertarTiemposMuertos(dbHelper, idFicha);

                            //enviar correo
                            //Crear lista de valores
                            List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado)
                            };
                            string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                            //lista de valores para resumen
                            List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                            string tablaResumen = GenerarTablaResumen(valoresResumen);
                            string tablaGenerada = GenerarTablaValores(valores);
                            // Generar tabla de tiempos muertos mecánicos
                            string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                            string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                            string resultado = GenerarTablaDetallesOP();
                            string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>
                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                            cuerpoHtml1 += cuerpoHtml2;
                            conf_email(cuerpoHtml1);
                            //fin correo
                            MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            cb_Area.SelectedIndex = -1;
                            reiniciarCampos();
                            cb_Area.Focus();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 6)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 
                    && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) 
                    && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) 
                    && !string.IsNullOrEmpty(Txt_6.Text) && !string.IsNullOrEmpty(Txt_7.Text) && !string.IsNullOrEmpty(cb_jefe_turno.Text) && ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                {
                    //Polvos
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = System.Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = System.Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal KgEntrada = System.Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = System.Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = System.Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = System.Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = System.Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = System.Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                    decimal polvo_colector = System.Convert.ToDecimal(Txt_6.Text);
                    decimal Granulo = System.Convert.ToDecimal(Txt_7.Text);
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                    string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                    if (editar)
                    {
                        idUsuarioActual = id_supervisor_global;
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        polvo_colector, Granulo, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);

                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Polvo de colector", polvo_colector.ToString()),
                                new KeyValuePair<string, string>("Granulo", Granulo.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado)
                            };
                        string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";
                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null,
                            KgEntrada, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                            polvo_colector, Granulo, 0, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, area, meta, 0, idUserSeleccionado);

                        if (idFicha > 0)
                        {
                            // Insertar en tablas relacionadas
                            InsertarTiemposMuertos(dbHelper, idFicha);

                            //enviar correo
                            //Crear lista de valores
                            List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", KgEntrada.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Polvo de colector", polvo_colector.ToString()),
                                new KeyValuePair<string, string>("Granulo", Granulo.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado)
                            };

                            string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                            //lista de valores para resumen
                            List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                            string tablaResumen = GenerarTablaResumen(valoresResumen);
                            string tablaGenerada = GenerarTablaValores(valores);
                            // Generar tabla de tiempos muertos mecánicos
                            string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                            string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                            string resultado = GenerarTablaDetallesOP();
                            string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                       <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                            cuerpoHtml1 += cuerpoHtml2;
                            conf_email(cuerpoHtml1);
                            //fin correo

                            MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            cb_Area.SelectedIndex = -1;
                            reiniciarCampos();
                            cb_Area.Focus();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (cb_Area.SelectedIndex == 8)
            {
                if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text) && !string.IsNullOrEmpty(txt_Tiempo_energia.Text) && cb_Turno.SelectedIndex != -1 
                    && Mask_txt_hr1.Text != "  :" && Mask_txt_hr2.Text != "  :" && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) 
                    && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_5.Text) 
                    && !string.IsNullOrEmpty(Txt_6.Text) && !string.IsNullOrEmpty(Txt_7.Text) && !string.IsNullOrEmpty(Txt_8.Text) 
                    && !string.IsNullOrEmpty(cb_jefe_turno.Text) && ValidarTodasLasFilas_operativo()&& ValidarTodasLasFilas_mecanico())
                {
                    //Revolturas
                    int idUsuarioActual = id_user;
                    // Obtener datos de los TextBox
                    DateTime fecha = dtp1.Value; // Tu MetroDateTime
                    int turno = System.Convert.ToInt32(cb_Turno.Text);
                    string op = cb_OP.Text;
                    decimal meta_kg = System.Convert.ToDecimal(Txt_Read_2.Text);
                    string textoPorcent_Logrado = Txt_Read_4.Text.Replace("%", "").Trim();
                    decimal Porcent_Logrado = System.Convert.ToDecimal(textoPorcent_Logrado) / 100m;
                    decimal Pz_producidas = System.Convert.ToDecimal(Txt_1.Text);
                    decimal KgProductoTerminado = System.Convert.ToDecimal(Txt_2.Text);
                    decimal KgFueraEspec = System.Convert.ToDecimal(Txt_3.Text);
                    decimal Merma = System.Convert.ToDecimal(Txt_4.Text);
                    int PersonalOpe = System.Convert.ToInt32(Txt_5.Text);
                    decimal hr_programadas = System.Convert.ToDecimal(Txt_Read_1.Text);
                    decimal hr_efec = System.Convert.ToDecimal(Txt_Read_3.Text);
                    // Conversión DIRECTA a TimeSpan desde los MaskedTextBox
                    TimeSpan hrInicio = TimeSpan.Parse(Mask_txt_hr1.Text);
                    TimeSpan hrFin = TimeSpan.Parse(Mask_txt_hr2.Text);
                    string area = cb_Area.Text;
                    decimal kg_entrada = System.Convert.ToDecimal(Txt_Read_5.Text);
                    decimal bobina_entrada = System.Convert.ToDecimal(Txt_6.Text);
                    decimal bobina_utilizada = System.Convert.ToDecimal(Txt_7.Text);
                    decimal bobina_merma = System.Convert.ToDecimal(Txt_8.Text);
                    decimal meta = System.Convert.ToDecimal(Txt_meta.Text);
                    string fecha_formateada = fecha.ToString("dd/MM/yyyy");

                    if (editar)
                    {
                        idUsuarioActual = id_supervisor_global;
                        updateFicha(dbHelper, idUsuarioActual, fecha, turno, op, null,
                        Pz_producidas, KgProductoTerminado, KgFueraEspec, Merma, 0, Porcent_Logrado,
                        kg_entrada, bobina_entrada, bobina_utilizada, hrInicio, hrFin, PersonalOpe, hr_programadas, hr_efec, meta_kg, 
                        null, meta, bobina_merma, idUserSeleccionado);

                        //enviar correo
                        //Crear lista de valores
                        List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Piezas Producidas", Pz_producidas.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Bobina Kg Entrada", bobina_entrada.ToString()),
                                new KeyValuePair<string, string>("Bobina Utilizada", bobina_utilizada.ToString()),
                                new KeyValuePair<string, string>("Bobina Merma", bobina_merma.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", kg_entrada.ToString())
                            };
                        string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                        //lista de valores para resumen
                        List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                        string tablaResumen = GenerarTablaResumen(valoresResumen);
                        string tablaGenerada = GenerarTablaValores(valores);
                        // Generar tabla de tiempos muertos mecánicos
                        string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                        string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                        string resultado = GenerarTablaDetallesOP();
                        string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>
                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                        cuerpoHtml1 += cuerpoHtml2;
                        conf_email(cuerpoHtml1);
                        //fin correo
                    }
                    else
                    {
                        // Insertar en tabla Ficha y obtener el ID_Ficha generado
                        int idFicha = InsertarFichaYRetornarID(dbHelper, idUsuarioActual, fecha, turno, op, null, Pz_producidas, KgProductoTerminado, 
                            KgFueraEspec, Merma, 0, Porcent_Logrado,kg_entrada, bobina_entrada, bobina_utilizada, hrInicio, hrFin, PersonalOpe, 
                            hr_programadas, hr_efec, meta_kg, area, meta, bobina_merma, idUserSeleccionado);
                        if (idFicha > 0)
                        {
                            // Insertar en tablas relacionadas
                            InsertarTiemposMuertos(dbHelper, idFicha);
                            //enviar correo
                            //Crear lista de valores
                            List<KeyValuePair<string, string>> valores = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Piezas Producidas", Pz_producidas.ToString()),
                                new KeyValuePair<string, string>("Kg Producto Terminado", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg Secos Fuera de Espec.", KgFueraEspec.ToString()),
                                new KeyValuePair<string, string>("Merma en Kg", Merma.ToString()),
                                new KeyValuePair<string, string>("Personal Operativo", PersonalOpe.ToString()),
                                new KeyValuePair<string, string>("Bobina Kg Entrada", bobina_entrada.ToString()),
                                new KeyValuePair<string, string>("Bobina Utilizada", bobina_utilizada.ToString()),
                                new KeyValuePair<string, string>("Bobina Merma", bobina_merma.ToString()),
                                new KeyValuePair<string, string>("Horas Progamadas", hr_programadas.ToString()),
                                new KeyValuePair<string, string>("Kg Meta Programada", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Horas Efectivas", hr_efec.ToString()),
                                new KeyValuePair<string, string>("Logro de Planeación (%)", textoPorcent_Logrado),
                                new KeyValuePair<string, string>("Kg entrada (proceso)", kg_entrada.ToString())
                            };
                            string Kgproducidosxhr = (KgProductoTerminado / hr_programadas).ToString("0.##");
                            //lista de valores para resumen
                            List<KeyValuePair<string, string>> valoresResumen = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Meta por Hora", meta.ToString()),
                                new KeyValuePair<string, string>("Meta programada en turno", meta_kg.ToString()),
                                new KeyValuePair<string, string>("Kg producidos en turno", KgProductoTerminado.ToString()),
                                new KeyValuePair<string, string>("Kg producidos por hora", Kgproducidosxhr)
                            };

                            string tablaResumen = GenerarTablaResumen(valoresResumen);
                            string tablaGenerada = GenerarTablaValores(valores);
                            // Generar tabla de tiempos muertos mecánicos
                            string tablaTiemposMuertos = GenerarTablaTiemposMuertosMecanicos();
                            string tablaTiemposMuertos_Operativo = GenerarTablaTiemposOperativos();
                            string resultado = GenerarTablaDetallesOP();
                            string cuerpoHtml2 = $@"
                                        <strong>{turno.ToString()}</strong>, inició su turno a las 
                                        <strong>{hrInicio.ToString()}</strong> y lo concluyó a las 
                                        <strong>{hrFin.ToString()}</strong>, el día 
                                        <strong>{fecha_formateada}</strong>.
                                    </p>

                                    <p>
                                        Durante el turno, se registró información en el área 
                                        <strong>{cb_Area.Text}</strong>, correspondiente al 
                                        <strong>{op}</strong>. La siguiente tabla muestra el resumen del turno:
                                    </p>
                                    {tablaResumen} 
                                    <br>
                                    <p>
                                        Y la siguiente tabla muestra los detalles registrados durante el turno:
                                    </p>

                                    {tablaGenerada}
                                    <br>
                                    <p>
                                        Además, se registraron los siguientes tiempos muertos mecánicos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos}
                                    <p>
                                        Y se registraron los siguientes tiempos muertos operativos durante el turno:
                                    </p>
                                    {tablaTiemposMuertos_Operativo}
                                    <p>
                                    Del cual los detalles del OP son los siguientes:
                                    </p>
                                        {resultado}     
                                    </br>  
                                    <p>        
                                        Agradecemos su atención. Este mensaje ha sido generado automáticamente por el Sistema Tablero.
                                    </p>

                                    <p style='margin-top: 20px;'>
                                        Atentamente,<br>
                                        <strong>Sistema Tablero</strong><br>
                                        Departamento de Sistemas
                                    </p>

                                </div>

                            </body>
                            </html>";

                            cuerpoHtml1 += cuerpoHtml2;
                            conf_email(cuerpoHtml1);
                            //fin correo
                            MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                                "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            cb_Area.SelectedIndex = -1;
                            reiniciarCampos();
                            cb_Area.Focus();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "Error al guardar datos",
                                                                "Operación fallida", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos requeridos.",
                                                        "Error de llenado", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (string.IsNullOrEmpty(Tablero.Properties.Settings.Default.ServerSMTP))
            {
                MetroFramework.MetroMessageBox.Show(this, "Complete la configuración SMTP antes de enviar el correo electronico. Solo el Administrador puede realizar esta configuración.", "SMTP pendiente", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void updateFicha(DatabaseHelper dbHelper, int idUsuario, DateTime fecha,
        int turno, string var1, string var2, decimal var3, decimal var4,
        decimal var5, decimal var6, decimal var7, decimal var8,
        decimal var9, decimal var10, decimal var11,
        TimeSpan hrInicio, TimeSpan hrFin, int personal_O, decimal hr_programadas, decimal hr_efectivas,
        decimal meta_kg, string area_f, decimal metaHr, decimal var12, int jefe)
        {
            DatabaseHelper dbHelper2 = new DatabaseHelper(connectionString);
            NpgsqlParameter[] parameters = new NpgsqlParameter[] { };
            string queryInsertUpdate = string.Empty;
            int result = 0;
            // Convertir el ID a entero ANTES de crear el parámetro
            int idFicha = System.Convert.ToInt32(id_global_ficha);
            if (cb_Area.SelectedIndex == 0)
            {
                // actualizar

                queryInsertUpdate = "UPDATE public.\"Ficha\" SET \"ID_user\" = @id_user, \"Fecha\" = @Fecha, \"Turno\" = @Turno, \"OP\" = @OP," +
                    " \"Kg_enter_proceso\" = @Kg_enter_proceso, \"kg_frescos_enter_se\" = @kg_frescos_enter_se, \"Merma_canica\" = @Merma_canica, \"Merma_podrido\" = @Merma_podrido," +
                    " \"Merma_tina\" = @Merma_tina, \"Merma_piso\" = @Merma_piso, \"Merma_canaletas\" = @Merma_canaletas, \"Merma_lavado_bandas\" = @Merma_lavado_bandas," +
                    " \"Cascara_carrete\" = @Cascara_carrete, \"Hr_inicio\" = @Hr_inicio, \"Hr_fin\" = @Hr_fin, \"Hr_programadas\" = @Hr_programadas, \"Personal_Operativo\" = @Personal_Operativo," +
                    " \"Hr_efectivas\" = @Hr_efectivas, \"Kg_meta\" = @Kg_meta, \"MetaHr\" = @MetaHr, \"ID_Jefe\" = @Jefe WHERE \"ID_Ficha\" = @ID_Ficha;";

                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@Turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Merma_canica", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_podrido", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@merma_tina", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@Merma_piso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Merma_canaletas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Merma_lavado_bandas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Cascara_carrete", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@Hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 1)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @op, ""Kg_prod_seco"" = @Kg_prod_seco,
                                    ""Merma_kg"" = @Merma_kg_seco, ""Kg_fuera_espec"" = @Kg_fuera_spec, ""Kg_resecar"" = @Kg_resecar, ""Personal_Operativo"" = @Personal_Operativo, 
                                    ""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, ""porcent_cump_meta"" = @Porcent_cumplimiento, ""Kg_meta"" = @Kg_meta,
                                    ""Relacion_Fr_seco"" = @Relacion_fresco_seco, ""FTT"" = @FTT, ""Hr_inicio"" = @hr_inicio, ""Hr_fin"" = @hr_fin, ""MetaHr"" = @MetaHr, 
                                    ""kg_frescos_enter_se"" = @kg_frescos_enter_se, ""ID_Jefe"" = @Jefe WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_prod_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Kg_fuera_spec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Merma_kg_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Kg_resecar", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Porcent_cumplimiento", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Relacion_fresco_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@FTT", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe}
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 2)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Kg_meta"" = @Kg_meta, 
                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, ""Kg_fuera_espec"" = @Kg_fuera_espec, 
                    ""Merma_kg"" = @Merma_kg,""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, ""Personal_Operativo"" = @Personal_Operativo, 
                    ""porcent_aumento_hum"" = @porcent_aumento_hum, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""MetaHr"" = @MetaHr, ""ID_Jefe"" = @Jefe WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@porcent_aumento_hum", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 7)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Kg_meta"" = @Kg_meta, 
                                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, 
                                    ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg,""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, 
                                    ""Personal_Operativo"" = @Personal_Operativo, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""MetaHr"" = @MetaHr, ""ID_Jefe"" = @Jefe WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 4)
            {

                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Proceso"" = @Proceso, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @op, ""Kg_meta"" = @Kg_meta, 
                                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, 
                                    ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg,""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, 
                                    ""Personal_Operativo"" = @Personal_Operativo, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""MetaHr"" = @MetaHr, ""ID_Jefe"" = @Jefe WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@Proceso", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 6)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Kg_meta"" = @Kg_meta, 
                                    ""porcent_cump_meta"" = @porcent_cump_meta, ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Kg_prod_term"" = @Kg_prod_term, 
                                    ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg, ""Hr_programadas"" = @Hr_programadas, ""Hr_efectivas"" = @Hr_efectivas, 
                                    ""Personal_Operativo"" = @Personal_Operativo, ""Hr_inicio"" = @Hr_inicio, ""Hr_fin"" = @Hr_fin, ""Polvo_colector"" = @Polvo_Colector, 
                                    ""Granulo"" = @Granulo, ""MetaHr"" = @MetaHr, ""ID_Jefe"" = @Jefe WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Polvo_Colector", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Granulo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }
            if (cb_Area.SelectedIndex == 8)
            {
                queryInsertUpdate = @"UPDATE public.""Ficha"" SET ""ID_user"" = @id_user, ""Fecha"" = @fecha, ""Turno"" = @turno, ""OP"" = @OP, ""Pz_prod"" = @Pz_prod, 
                    ""Kg_prod_term"" = @Kg_prod_term, ""Kg_fuera_espec"" = @Kg_fuera_espec, ""Merma_kg"" = @Merma_kg, ""Kg_meta"" = @Kg_meta, ""porcent_cump_meta"" = @porcent_cump_meta, 
                    ""Kg_enter_proceso"" = @Kg_enter_proceso, ""Bobina_kg_enter"" = @Bobina_kg_enter, ""Bobina_utilizada"" = @Bobina_utilizada, ""Hr_inicio"" = @Hr_inicio, 
                    ""Hr_fin"" = @Hr_fin, ""Personal_Operativo"" = @Personal_Operativo, ""Hr_programadas"" = @Hr_programadas, ""Bobina_merma"" = @Bobina_merma, 
                    ""Hr_efectivas"" = @Hr_efectivas, ""MetaHr"" = @MetaHr, ""ID_Jefe"" = @Jefe WHERE ""ID_Ficha"" = @ID_Ficha;";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Pz_prod", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Bobina_kg_enter", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Bobina_utilizada", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Bobina_merma", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var12 },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer){Value = idFicha},
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
                result = dbHelper2.ExecuteNonQuery(queryInsertUpdate, parameters);
            }

            // Actualizar los tiempos muertos asociados
            UpdateTiemposMuertos(dbHelper);
            if (result > 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "Datos guardados correctamente",
                                                            "Operación exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cb_Area.SelectedIndex = -1;
                reiniciarCampos();
                cb_Area.Focus();
                if (_editarForm == null || _editarForm.IsDisposed)
                {
                    editar = false; // Reiniciar el estado de edición
                }
                else
                {
                    // Traer al frente el formulario
                    _editarForm.BringToFront(); // Opción 1
                                                // _editarForm.Activate(); // Opción 2 (también funciona)

                    // Opcional: También puedes enfocarlo
                    _editarForm.Focus();
                }
                id_global_ficha = string.Empty; // Limpiar el ID global
                if (nivel_user == "Super Administrador")
                {
                    lbl_no_emp2.Text = numero_empleado_admin;
                    lbl_nom2.Text = nombre_admin;
                    cb_supervisor_turno.Visible = false;
                }
            }
        }

        private int InsertarFichaYRetornarID(DatabaseHelper dbHelper, int idUsuario, DateTime fecha,
        int turno, string var1, string var2, decimal var3, decimal var4,
        decimal var5, decimal var6, decimal var7, decimal var8,
        decimal var9, decimal var10, decimal var11,
        TimeSpan hrInicio, TimeSpan hrFin, int personal_O, decimal hr_programadas, decimal hr_efectivas,
        decimal meta_kg, string area_f, decimal metaHr, decimal var12, int jefe)
        {
            string query = String.Empty;
            NpgsqlParameter[] parameters = new NpgsqlParameter[] { };
            if (cb_Area.SelectedIndex == 0)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""Lote"", ""OP"",
                    ""Kg_enter_proceso"", ""kg_frescos_enter_se"", ""Merma_canica"",
                    ""Merma_podrido"", ""Merma_tina"", ""Merma_piso"", ""Merma_canaletas"",
                    ""Merma_lavado_bandas"", ""Cascara_carrete"", ""Hr_inicio"", ""Hr_fin"", 
                    ""Hr_programadas"", ""Personal_Operativo"", ""Hr_efectivas"", ""Kg_meta"", ""Area"", ""MetaHr"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @fecha, @turno, @lote, @op,
                    @kg_enter_proceso, @kg_frescos_enter_se, @merma_canica,
                    @merma_podrido, @merma_tina, @merma_piso, @merma_canaletas,
                    @merma_lavado_bandas, @cascara_carrete, @hr_inicio, @hr_fin, @hr_prog, @Personal_Op, @Hr_efec, @Kg_meta, @Area, @MetaHr, @Jefe
                ) RETURNING ""ID_Ficha"";";

                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@lote", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@merma_canica", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@merma_podrido", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@merma_tina", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@merma_piso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@merma_canaletas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@merma_lavado_bandas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@cascara_carrete", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@hr_prog", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Personal_Op", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_efec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f ?? (object)DBNull.Value },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            if (cb_Area.SelectedIndex == 1)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Lote"", ""Kg_prod_seco"", ""Merma_kg"", ""Kg_fuera_espec"", ""Kg_resecar"", 
                    ""Personal_Operativo"", ""Hr_programadas"", ""Hr_efectivas"", ""porcent_cump_meta"", ""Kg_meta"",
                    ""Relacion_Fr_seco"", ""FTT"", ""Hr_inicio"", ""Hr_fin"", ""Area"", ""MetaHr"", ""kg_frescos_enter_se"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @fecha, @turno, @op, @Lote, @Kg_prod_seco, @Merma_kg_seco, @Kg_fuera_spec,
                    @Kg_resecar, @Personal_Operativo, @Hr_programadas,
                    @Hr_efectivas, @Porcent_cumplimiento, @Kg_meta,
                    @Relacion_fresco_seco, @FTT, @hr_inicio, @hr_fin, @Area, @MetaHr, @kg_frescos_enter_se, @Jefe
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@op", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Lote", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_prod_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Kg_fuera_spec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Merma_kg_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Kg_resecar", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Integer) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Porcent_cumplimiento", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Relacion_fresco_seco", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var7 },
                    new NpgsqlParameter("@FTT", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@kg_frescos_enter_se", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            if (cb_Area.SelectedIndex == 2)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", 
                    ""Kg_prod_term"", ""Kg_fuera_espec"", ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", 
                    ""porcent_aumento_hum"", ""Hr_inicio"", ""Hr_fin"", ""Area"", ""MetaHr"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta,
                    @Kg_enter_proceso, @Kg_prod_term, @Kg_fuera_espec,
                    @Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo,
                    @porcent_aumento_hum, @Hr_inicio, @Hr_fin, @Area, @MetaHr, @Jefe
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@porcent_aumento_hum", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            if (cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 5)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""Area"", ""MetaHr"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Kg_prod_term, 
                    @Kg_fuera_espec,@Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo, @Hr_inicio, @Hr_fin, @Area, @MetaHr, @Jefe
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            if (cb_Area.SelectedIndex == 4)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Proceso"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", ""Area"", 
                    ""MetaHr"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @Proceso, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Kg_prod_term, 
                    @Kg_fuera_espec,@Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo, @Hr_inicio, @Hr_fin, @Area, @MetaHr, @Jefe
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@Proceso", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var2 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            if (cb_Area.SelectedIndex == 6)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""porcent_cump_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", ""Kg_fuera_espec"", 
                    ""Merma_kg"",""Hr_programadas"", ""Hr_efectivas"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", ""Polvo_colector"", ""Granulo"", 
                    ""Area"", ""MetaHr"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Kg_prod_term, 
                    @Kg_fuera_espec,@Merma_kg, @Hr_programadas, @Hr_efectivas, @Personal_Operativo, @Hr_inicio, @Hr_fin, @Polvo_Colector, @Granulo, @Area, 
                    @MetaHr, @Jefe
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Polvo_Colector", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Granulo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            if (cb_Area.SelectedIndex == 8)
            {
                query = @"INSERT INTO public.""Ficha"" (
                    ""ID_user"", ""Fecha"", ""Turno"", ""OP"", ""Pz_prod"", ""Kg_prod_term"", ""Kg_fuera_espec"", ""Merma_kg"", ""Kg_meta"", ""porcent_cump_meta"", 
                    ""Kg_enter_proceso"", ""Bobina_kg_enter"", ""Bobina_utilizada"", ""Hr_inicio"", ""Hr_fin"", ""Personal_Operativo"", ""Hr_programadas"", ""Bobina_merma"", 
                    ""Hr_efectivas"", ""Area"", ""MetaHr"", ""ID_Jefe""
                ) VALUES (
                    @id_user, @fecha, @turno, @OP, @Pz_prod, @Kg_prod_term, @Kg_fuera_espec, @Merma_kg, @Kg_meta, @porcent_cump_meta, @Kg_enter_proceso, @Bobina_kg_enter,  
                    @Bobina_utilizada, @Hr_inicio, @Hr_fin, @Personal_Operativo, @Hr_programadas, @Bobina_merma, @Hr_efectivas, @Area, @MetaHr, @Jefe
                ) RETURNING ""ID_Ficha"";";
                parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_user", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idUsuario },
                    new NpgsqlParameter("@fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                    new NpgsqlParameter("@turno", NpgsqlTypes.NpgsqlDbType.Integer) { Value = turno },
                    new NpgsqlParameter("@OP", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
                    new NpgsqlParameter("@Pz_prod", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var3 },
                    new NpgsqlParameter("@Kg_prod_term", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var4 },
                    new NpgsqlParameter("@Kg_fuera_espec", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var5 },
                    new NpgsqlParameter("@Merma_kg", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var6 },
                    new NpgsqlParameter("@Kg_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = meta_kg },
                    new NpgsqlParameter("@porcent_cump_meta", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var8 },
                    new NpgsqlParameter("@Kg_enter_proceso", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var9 },
                    new NpgsqlParameter("@Bobina_kg_enter", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var10 },
                    new NpgsqlParameter("@Bobina_utilizada", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var11 },
                    new NpgsqlParameter("@hr_inicio", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrInicio },
                    new NpgsqlParameter("@hr_fin", NpgsqlTypes.NpgsqlDbType.Time) { Value = hrFin },
                    new NpgsqlParameter("@Personal_Operativo", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = personal_O },
                    new NpgsqlParameter("@Hr_programadas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_programadas },
                    new NpgsqlParameter("@Bobina_merma", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = var12 },
                    new NpgsqlParameter("@Hr_efectivas", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = hr_efectivas },
                    new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = area_f },
                    new NpgsqlParameter("@MetaHr", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = metaHr },
                    new NpgsqlParameter("@Jefe", NpgsqlTypes.NpgsqlDbType.Integer) { Value = jefe }
                };
            }
            return dbHelper.ExecuteScalarInt(query, parameters);
        }

        private void UpdateTiemposMuertos(DatabaseHelper dbHelper)
        {
            // Insertar tiempos muertos mecánicos desde DataGridView
            UpdateTiemposMuertosMecanicos(dbHelper);

            // Insertar tiempos muertos operativos desde DataGridView
            UpdateTiemposMuertosOperativos(dbHelper);

            // Insertar tiempo muerto comida desde TextBox
            UpdateTiempoMuertoComida(dbHelper);

            // Insertar tiempo muerto energía desde TextBox
            UpdateTiempoMuertoEnergia(dbHelper);
        }
        private void UpdateTiemposMuertosMecanicos(DatabaseHelper dbHelper)
        {
            int id_ficha_int = System.Convert.ToInt32(id_global_ficha);

            // Primero eliminamos todos los registros existentes para esta ficha
            string deleteQuery = @"DELETE FROM public.""Tiempo_muerto_Mecanico""
                         WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] deleteParams = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int }
            };

            dbHelper.ExecuteNonQuery(deleteQuery, deleteParams);

            InsertarTiemposMuertosMecanicos(dbHelper, id_ficha_int);
        }
        private void UpdateTiemposMuertosOperativos(DatabaseHelper dbHelper)
        {
            int id_ficha_int = System.Convert.ToInt32(id_global_ficha);

            // Primero eliminamos todos los registros existentes para esta ficha
            string deleteQuery = @"DELETE FROM public.""Tiempo_Muerto_Operativo"" 
                         WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] deleteParams = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int }
            };

            dbHelper.ExecuteNonQuery(deleteQuery, deleteParams);

            // Luego insertamos todos los registros del DataGridView
            InsertarTiemposMuertosOperativos(dbHelper, id_ficha_int);
        }
        private void UpdateTiempoMuertoComida(DatabaseHelper dbHelper)
        {
            int id_ficha_int = System.Convert.ToInt32(id_global_ficha);
            if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text))
            {
                decimal minutosDetenidos = System.Convert.ToDecimal(txt_Tiempo_comida.Text);

                string query = @"UPDATE public.""Tiempo_Muerto_Comida"" SET ""Minutos_Detenidos"" = @minutos_detenidos WHERE ""ID_Ficha"" = @id_ficha;";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }
        private void UpdateTiempoMuertoEnergia(DatabaseHelper dbHelper)
        {
            int id_ficha_int = System.Convert.ToInt32(id_global_ficha);
            if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text))
            {
                decimal minutosDetenidos = System.Convert.ToDecimal(txt_Tiempo_energia.Text);

                string query = @"UPDATE public.""Tiempo_Muerto_Energia"" SET ""Minutos_Detenidos"" = @minutos_detenidos WHERE ""ID_Ficha"" = @id_ficha;";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = id_ficha_int },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }
        private void InsertarTiemposMuertos(DatabaseHelper dbHelper, int idFicha)
        {
            // Insertar tiempos muertos mecánicos desde DataGridView
            InsertarTiemposMuertosMecanicos(dbHelper, idFicha);

            // Insertar tiempos muertos operativos desde DataGridView
            InsertarTiemposMuertosOperativos(dbHelper, idFicha);

            // Insertar tiempo muerto comida desde TextBox
            InsertarTiempoMuertoComida(dbHelper, idFicha);

            // Insertar tiempo muerto energía desde TextBox
            InsertarTiempoMuertoEnergia(dbHelper, idFicha);
        }

        // Insertar en Tiempo_muerto_Mecanico desde DataGridView
        private void InsertarTiemposMuertosMecanicos(DatabaseHelper dbHelper, int idFicha)
        {
            valores_mecanico.Clear(); // Limpiar la lista antes de agregar nuevos valores
            foreach (DataGridViewRow row in dgv_mecanico.Rows)
            {
                if (!row.IsNewRow && row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    decimal minDetenidos = System.Convert.ToDecimal(row.Cells[0].Value);
                    string tipo = row.Cells[1].Value.ToString();
                    string motivos = row.Cells[2].Value.ToString();

                    string query = @"INSERT INTO public.""Tiempo_muerto_Mecanico"" (
                            ""ID_Ficha"", ""Min_Detenidos"", ""Motivos"", ""Tipo""
                        ) VALUES (@id_ficha, @min_detenidos, @motivos, @Tipo);";

                    NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                        new NpgsqlParameter("@min_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minDetenidos },
                        new NpgsqlParameter("@motivos", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = motivos ?? (object)DBNull.Value },
                        new NpgsqlParameter("@Tipo", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = tipo ?? (object)DBNull.Value }
                    };

                    dbHelper.ExecuteNonQuery(query, parameters);
                    valores_mecanico.Add((row.Cells[0].Value.ToString(), tipo, motivos));
                }
            }
        }

        // Insertar en Tiempo_Muerto_Operativo desde DataGridView
        private void InsertarTiemposMuertosOperativos(DatabaseHelper dbHelper, int idFicha)
        {
            valores_operativos.Clear(); // Limpiar la lista antes de agregar nuevos valores
            foreach (DataGridViewRow row in dgv_operativo.Rows)
            {
                if (!row.IsNewRow && row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    decimal minDetenidos = System.Convert.ToDecimal(row.Cells[0].Value);
                    string tipo = row.Cells[1].Value.ToString();
                    string motivos = row.Cells[2].Value.ToString();

                    string query = @"INSERT INTO public.""Tiempo_Muerto_Operativo"" (
                            ""ID_Ficha"", ""Min_Detenidos"", ""Motivos"", ""Tipo""
                        ) VALUES (@id_ficha, @min_detenidos, @motivos, @Tipo);";

                    NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                        new NpgsqlParameter("@min_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minDetenidos },
                        new NpgsqlParameter("@motivos", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = motivos ?? (object)DBNull.Value },
                        new NpgsqlParameter("@Tipo", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = tipo ?? (object)DBNull.Value }
                    };

                    dbHelper.ExecuteNonQuery(query, parameters);
                    valores_operativos.Add((row.Cells[0].Value.ToString(), tipo, motivos));
                }
            }
        }

        // Insertar en Tiempo_Muerto_Comida desde TextBox
        private void InsertarTiempoMuertoComida(DatabaseHelper dbHelper, int idFicha)
        {
            if (!string.IsNullOrEmpty(txt_Tiempo_comida.Text))
            {
                decimal minutosDetenidos = System.Convert.ToDecimal(txt_Tiempo_comida.Text);

                string query = @"INSERT INTO public.""Tiempo_Muerto_Comida"" (
                        ""ID_Ficha"", ""Minutos_Detenidos""
                    ) VALUES (@id_ficha, @minutos_detenidos);";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }

        // Insertar en Tiempo_Muerto_Energia desde TextBox
        private void InsertarTiempoMuertoEnergia(DatabaseHelper dbHelper, int idFicha)
        {
            if (!string.IsNullOrEmpty(txt_Tiempo_energia.Text))
            {
                decimal minutosDetenidos = System.Convert.ToDecimal(txt_Tiempo_energia.Text);

                string query = @"INSERT INTO public.""Tiempo_Muerto_Energia"" (
                        ""ID_Ficha"", ""Minutos_Detenidos""
                    ) VALUES (@id_ficha, @minutos_detenidos);";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = idFicha },
                    new NpgsqlParameter("@minutos_detenidos", NpgsqlTypes.NpgsqlDbType.Numeric) { Value = minutosDetenidos }
                };

                dbHelper.ExecuteNonQuery(query, parameters);
            }
        }

        private void CalcularTotalMecanico()
        {
            decimal total = 0;

            foreach (DataGridViewRow row in dgv_mecanico.Rows)
            {
                if (row.IsNewRow) continue; // ignora la fila en blanco para agregar

                if (row.Cells[0].Value != null &&
                    decimal.TryParse(row.Cells[0].Value.ToString(), out decimal valor))
                {
                    total += valor;
                }
            }
            txt_TM_mecanico.Text = total.ToString("0"); // formato con 2 decimales
        }

        private void CalcularTotaloperativo()
        {
            decimal total = 0;

            foreach (DataGridViewRow row in dgv_operativo.Rows)
            {
                if (row.IsNewRow) continue; // ignora la fila en blanco para agregar

                if (row.Cells[0].Value != null &&
                    decimal.TryParse(row.Cells[0].Value.ToString(), out decimal valor))
                {
                    total += valor;
                }
            }
            txt_TM_operativo.Text = total.ToString("0"); // formato con 2 decimales
        }

        private void dgv_mecanico_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            CalcularTotalMecanico();
        }

        private void dgv_mecanico_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            CalcularTotalMecanico();
        }

        private void dgv_mecanico_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            CalcularTotalMecanico();
        }
        // MÉTODO CellValueChanged COMBINADO
        private void dgv_operativo_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Evitar errores
            if (e.RowIndex < 0) return;

            // Si cambió la columna Tipo
            if (dgv_operativo.Columns[e.ColumnIndex].Name == "columTipo")
            {
                DataGridViewComboBoxCell cmbMotivo =
                    dgv_operativo.Rows[e.RowIndex].Cells["columMotivo"] as DataGridViewComboBoxCell;

                if (cmbMotivo == null) return;

                cmbMotivo.Items.Clear();

                string tipo = dgv_operativo.Rows[e.RowIndex]
                               .Cells["columTipo"].Value?.ToString();

                if (tipo == "ALMACÉN")
                {
                    cmbMotivo.Items.Add("ESPERA DE PRODUCTO");
                    cmbMotivo.Items.Add("ESPERA DE INSUMOS");
                }
                else if (tipo == "CALIDAD")
                {
                    cmbMotivo.Items.Add("PRODUCTO SIN LIBERAR");
                    cmbMotivo.Items.Add("ESPERA POR AW");
                    cmbMotivo.Items.Add("DETENIDO POR PRODUCTO FUERA DE ESPECIFICACIÓN");
                    cmbMotivo.Items.Add("PRESENCIA DE MATERIA EXTRAÑA");
                }
                else if (tipo == "PRODUCCIÓN")
                {
                    cmbMotivo.Items.Add("CAMBIO DE LOTE");
                    cmbMotivo.Items.Add("ORGANIZACIÓN DE ARRANQUE");
                    cmbMotivo.Items.Add("ACOMODO DE PERSONAL");
                    cmbMotivo.Items.Add("LIMPIEZAS");
                    cmbMotivo.Items.Add("PREPARACIÓN DE ÁREA");
                }

                // Limpiar selección previa
                cmbMotivo.Value = null;
            }

           
            CalcularTotaloperativo();
        }

        private void dgv_operativo_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            CalcularTotaloperativo();
        }

        private void dgv_operativo_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            CalcularTotaloperativo();
        }

        private void Txt_3_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
            if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                porcentaje_logrado_planeacion();
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                calcular_meta_programada();
                porcentaje_logrado_planeacion();
            }
        }

        private void Txt_4_TextChanged(object sender, EventArgs e)
        {
            var tb = (RadTextBox)sender;
            string original = tb.Text;
            string saneado = SanitizeNumericText(original);

            if (saneado != original)
            {
                int sel = tb.SelectionStart;
                tb.Text = saneado;
                // Ajusta la posición del cursor (no exceder la longitud)
                tb.SelectionStart = Math.Min(sel, tb.Text.Length);
            }

            if (cb_Area.SelectedIndex == 0)
            {
                CalcularSuma();
            }
            if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_cumplimiento_metas();
                Ftt_metodo();
                Relacion_Fresco_seco();
            }
        }

        private void Txt_5_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                calcular_meta_programada();
                porcentaje_logrado_planeacion();
            }
        }

        private void Txt_6_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
        }


        private void Txt_7_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
        }

        private void Txt_8_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                CalcularSuma();
            }
        }

        private void txt_TM_mecanico_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }

        private void txt_TM_operativo_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
        }

        private void Txt_2_TextChanged(object sender, EventArgs e)
        {
            var tb = (RadTextBox)sender;
            string original = tb.Text;
            string saneado = SanitizeNumericText(original);

            if (saneado != original)
            {
                int sel = tb.SelectionStart;
                tb.Text = saneado;
                // Ajusta la posición del cursor (no exceder la longitud)
                tb.SelectionStart = Math.Min(sel, tb.Text.Length);
            }

            if (cb_Area.SelectedIndex == 0)
            {
                CalcularSuma();
            }
            if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_cumplimiento_metas();
                Ftt_metodo();
                Relacion_Fresco_seco();
            }
            if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion();
            }
            if (cb_Area.SelectedIndex == 2 && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text))
            {
                porcentaje_aumento_humedad();
            }
            if ((cb_Area.SelectedIndex == 8 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion_platinum();
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                calcular_meta_programada();
                porcentaje_logrado_planeacion();
            }

        }

        // Función que deja sólo dígitos y un solo punto (mantiene el primer punto)
        private string SanitizeNumericText(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            // Elimina todo excepto dígitos y puntos
            string onlyDigitsAndDots = Regex.Replace(input, @"[^0-9.]", "");

            int firstDot = onlyDigitsAndDots.IndexOf('.');
            if (firstDot >= 0)
            {
                // deja sólo el primer punto
                string before = onlyDotsAndTrim(onlyDigitsAndDots.Substring(0, firstDot + 1));
                string after = onlyDigitsAndDots.Substring(firstDot + 1).Replace(".", ""); // eliminar puntos siguientes
                return before + after;
            }
            else
            {
                return onlyDotsAndTrim(onlyDigitsAndDots);
            }

            // local helper
            string onlyDotsAndTrim(string s) => s; // ya está limpio en este flujo
        }
        private void cb_OP_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 0)
            {
                cb_Turno.Enabled = true;
            }
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 8)
            {
                cb_Turno.Enabled = true;
            }

            buscar_Meta_hr();
        }

        private void buscar_Meta_hr()
        {
            string valorBuscado = cb_OP.Text;

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            if (cb_Area.SelectedIndex == 0 || cb_Area.SelectedIndex == 1)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"No_box_hr\" FROM public.\"Deshidratado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["No_box_hr"].ToString();
                    Txt_meta.Text = resultado;

                    if (!string.IsNullOrEmpty(Txt_Read_1.Text))
                    {
                        calcular_meta_programada();
                    }
                }
            }
            if (cb_Area.SelectedIndex == 1)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Kg_seco_hr\" FROM public.\"Deshidratado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Kg_seco_hr"].ToString();
                    Txt_meta.Text = resultado;

                    if (!string.IsNullOrEmpty(Txt_Read_1.Text))
                    {
                        calcular_meta_programada();
                    }
                }
            }
            if (cb_Area.SelectedIndex == 2)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_kg_hr\" FROM public.\"Evaporado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_kg_hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 3)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_Kg_hr\" FROM public.\"Grind\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_Kg_hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 4)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_kg_hr_line\" FROM public.\"Inspeccion\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_kg_hr_line"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 5)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_kg_hr_line\" FROM public.\"Empacado\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_kg_hr_line"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 6)
            {
                // Consulta para buscar donde OP = valor_buscado
                // Obtiene la fecha actual
                DateTime hoy = DateTime.Now;

                // Define las fechas de inicio y fin del rango (solo cambia el año actual)
                DateTime inicio = new DateTime(hoy.Year, 5, 1);   // 1 de mayo
                DateTime fin = new DateTime(hoy.Year, 9, 30);     // 30 de septiembre

                string query;

                // Valida si la fecha actual está dentro del rango
                if (hoy >= inicio && hoy <= fin)
                {
                    query = "SELECT \"Meta_kg_hr_hum\" FROM public.\"Polvos\" WHERE \"OP\" = @valorBuscado;";
                }
                else
                {
                    query = "SELECT \"Meta_kg_hr_idon\" FROM public.\"Polvos\" WHERE \"OP\" = @valorBuscado;";
                }

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {

                    if (hoy >= inicio && hoy <= fin)
                    {
                        resultado = dt.Rows[0]["Meta_kg_hr_hum"].ToString();
                    }
                    else
                    {
                        resultado = dt.Rows[0]["Meta_kg_hr_idon"].ToString();
                    }
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 7)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_Kg_Hr\" FROM public.\"Revolturas\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_Kg_Hr"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
            if (cb_Area.SelectedIndex == 8)
            {
                // Consulta para buscar donde OP = valor_buscado
                string query = "SELECT \"Meta_Kg_Hr\", \"Kg_pz\" FROM public.\"Maquinas\" WHERE \"OP\" = @valorBuscado;";

                // Crear parámetro
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                };

                // Ejecutar consulta
                System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                // Variable string donde guardar el resultado
                string resultado = string.Empty;

                // Verificar si se encontraron resultados
                if (dt != null && dt.Rows.Count > 0)
                {
                    resultado = dt.Rows[0]["Meta_Kg_Hr"].ToString();
                    Kg_pz_str  = dt.Rows[0]["Kg_pz"].ToString();
                    Txt_meta.Text = resultado;
                }
            }
        }

        private void txt_Tiempo_comida_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
            if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_cumplimiento_metas();
                Ftt_metodo();
                Relacion_Fresco_seco();
            }
            if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion();
            }
            if (cb_Area.SelectedIndex == 2 && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text))
            {
                porcentaje_aumento_humedad();
            }
            if ((cb_Area.SelectedIndex == 8 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion_platinum();
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                calcular_meta_programada();
                porcentaje_logrado_planeacion();
            }
        }

        private void Mask_txt_hr1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                horaInicioText = Mask_txt_hr1.Text;
                // Convertir a DateTime
                horaInicio = DateTime.ParseExact(horaInicioText, "HH:mm", null);
                errorProvider1.SetError(Mask_txt_hr1, ""); // Borra el error si es válido
                btn_save_ficha.Enabled = true;
                calcular_turno();
                //if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
                //{
                //    calcular_meta_programada();
                //    porcentaje_logrado_planeacion();
                //}

                if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
                {
                    porcentaje_cumplimiento_metas();
                    Ftt_metodo();
                    Relacion_Fresco_seco();
                }
                if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
                {
                    porcentaje_logrado_planeacion();
                }
                if (cb_Area.SelectedIndex == 2 && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text))
                {
                    porcentaje_aumento_humedad();
                }
                if ((cb_Area.SelectedIndex == 8 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
                {
                    porcentaje_logrado_planeacion_platinum();
                }
                if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
                {
                    calcular_meta_programada();
                    porcentaje_logrado_planeacion();
                }
            }
            catch
            {
                errorProvider1.SetError(Mask_txt_hr1, "Formato de hora inválido. Asegúrese de usar HH:mm.");
                btn_save_ficha.Enabled = false;
            }
        }

        private void Mask_txt_hr2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                horaFinText = Mask_txt_hr2.Text;
                // Convertir a DateTime
                horaFin = DateTime.ParseExact(horaFinText, "HH:mm", null);
                errorProvider2.SetError(Mask_txt_hr2, ""); // Borra el error si es válido
                btn_save_ficha.Enabled = true;
                calcular_turno();
                //if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
                //{
                //    calcular_meta_programada();
                //    porcentaje_logrado_planeacion();
                //}

                if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
                {
                    porcentaje_cumplimiento_metas();
                    Ftt_metodo();
                    Relacion_Fresco_seco();
                }
                if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
                {
                    porcentaje_logrado_planeacion();
                }
                if (cb_Area.SelectedIndex == 2 && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text))
                {
                    porcentaje_aumento_humedad();
                }
                if ((cb_Area.SelectedIndex == 8 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
                {
                    porcentaje_logrado_planeacion_platinum();
                }
                if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
                {
                    calcular_meta_programada();
                    porcentaje_logrado_planeacion();
                }
            }
            catch
            {
                errorProvider2.SetError(Mask_txt_hr2, "Formato de hora inválido. Asegúrese de usar HH:mm.");
                btn_save_ficha.Enabled = false;
            }
        }

        private void txt_Tiempo_energia_TextChanged(object sender, EventArgs e)
        {
            calcular_turno();
            if (cb_Area.SelectedIndex == 1 && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_4.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_cumplimiento_metas();
                Ftt_metodo();
                Relacion_Fresco_seco();
            }
            if ((cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 6) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_3.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion();
            }
            if (cb_Area.SelectedIndex == 2 && !string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text))
            {
                porcentaje_aumento_humedad();
            }
            if ((cb_Area.SelectedIndex == 8 || cb_Area.SelectedIndex == 7) && !string.IsNullOrEmpty(Txt_2.Text) && !string.IsNullOrEmpty(Txt_Read_2.Text))
            {
                porcentaje_logrado_planeacion_platinum();
            }
            if (cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5)
            {
                calcular_meta_programada();
                porcentaje_logrado_planeacion();
            }
        }

        private void Txt_2_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_2_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void Txt_3_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_4_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_5_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_6_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_7_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_8_Validating(object sender, CancelEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string text = textBox.Text;

            // Regex que acepta "" o dígitos con un punto decimal opcional
            if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
            {
                MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true; // mantiene el foco en el control
            }
        }

        private void Txt_9_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir control de borrado (Backspace)
            if (e.KeyChar == (char)Keys.Back)
            {
                return;
            }

            // Permitir solo un punto decimal
            if (e.KeyChar == '.')
            {
                if ((sender as RadTextBox).Text.Contains("."))
                {
                    e.Handled = true; // Ya hay un punto → se bloquea
                }
                return;
            }

            // Permitir solo números
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Bloquear todo lo que no sea número
            }
        }

        private void Txt_10_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir control de borrado (Backspace)
            if (e.KeyChar == (char)Keys.Back)
            {
                return;
            }

            // Permitir solo un punto decimal
            if (e.KeyChar == '.')
            {
                if ((sender as RadTextBox).Text.Contains("."))
                {
                    e.Handled = true; // Ya hay un punto → se bloquea
                }
                return;
            }

            // Permitir solo números
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Bloquear todo lo que no sea número
            }
        }

        private void porcentaje_logrado_planeacion()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            if (double.TryParse(Txt_Read_2.Text, out double meta_kg) && double.TryParse(Txt_2.Text, out double Kg_Prod_Terminado) && double.TryParse(Txt_3.Text, out double kg_fuera_espec))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (meta_kg == 0)
                {
                    Txt_Read_4.Text = "";
                    return;
                }

                // Calcular la fórmula: ((Q2 - S2) / P2)
                double resultado = (Kg_Prod_Terminado - kg_fuera_espec) / meta_kg;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_4.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }
        private void porcentaje_logrado_planeacion_platinum()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            if (double.TryParse(Txt_Read_2.Text, out double meta_kg) && double.TryParse(Txt_2.Text, out double Kg_Prod_Terminado))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (meta_kg == 0)
                {
                    Txt_Read_4.Text = "";
                    return;
                }

                double resultado = Kg_Prod_Terminado / meta_kg;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_4.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }
        private void calcular_kg_entrada_proceso()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta para buscar donde OP = valor_buscado
            string query = "SELECT \"Kg_pz\" FROM public.\"Maquinas\" WHERE \"OP\" = @valorBuscado;";
            string valorBuscado = cb_OP.Text;
            string kg_pz_string;
            decimal kg_pz, pz_producidas, resultado;
            kg_pz = 0;

            // Crear parámetro
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
            };

            // Ejecutar consulta
            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            // Verificar si se encontraron resultados
            if (dt != null && dt.Rows.Count > 0)
            {
                kg_pz_string = dt.Rows[0]["Kg_pz"].ToString();
                kg_pz = System.Convert.ToDecimal(kg_pz_string);
            }
            pz_producidas = System.Convert.ToDecimal(Txt_1.Text);
            resultado = pz_producidas * kg_pz;

            Txt_Read_5.Text = resultado.ToString();
        }

        private void porcentaje_aumento_humedad()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            if (double.TryParse(Txt_1.Text, out double Kg_entrada) && double.TryParse(Txt_2.Text, out double Kg_Prod_Terminado))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (Kg_entrada == 0)
                {
                    Txt_Read_5.Text = "";
                    return;
                }

                // Calcular la fórmula: ((Q2 - S2) / P2)
                double resultado = (Kg_Prod_Terminado / Kg_entrada) - 1;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_5.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }

        private void porcentaje_cumplimiento_metas()
        {
            if (double.TryParse(Txt_2.Text, out double kg_prod_seco) && double.TryParse(Txt_4.Text, out double kg_fuera_espec) && double.TryParse(Txt_Read_2.Text, out double kg_secos_meta))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (kg_secos_meta == 0)
                {
                    Txt_Read_5.Text = "";
                    return;
                }

                // Calcular la fórmula: ((Q2 - S2) / P2)
                double resultado = (kg_prod_seco - kg_fuera_espec) / kg_secos_meta;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_5.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }

        private void Ftt_metodo()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            if (double.TryParse(Txt_2.Text, out double kg_prod_seco) && double.TryParse(Txt_4.Text, out double kg_fuera_espec))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (kg_prod_seco == 0)
                {
                    Txt_Read_8.Text = "";
                    return;
                }

                double resultado = (kg_prod_seco - kg_fuera_espec) / kg_prod_seco;

                // Aplicar la condición: si es mayor a 1 (100%), usar 1 (100%)
                if (resultado > 1.0)
                {
                    resultado = 1.0;
                }

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_8.Text = resultado.ToString("P2"); // Formato de porcentaje con 2 decimales
            }
        }

        private void Relacion_Fresco_seco()
        {
            // Obtener los valores de los TextBox y convertirlos a double

            if (double.TryParse(Txt_Read_4.Text, out double kg_frescos_enter_sec) && double.TryParse(Txt_2.Text, out double kg_prod_seco))
            {
                // Verificar que P2 no sea cero para evitar división por cero
                if (kg_prod_seco == 0)
                {
                    Txt_Read_7.Text = "";
                    return;
                }

                double resultado = kg_frescos_enter_sec / kg_prod_seco;

                // Convertir a porcentaje y mostrar en el TextBox de resultado
                Txt_Read_7.Text = resultado.ToString("0.##"); // Formato de porcentaje con 2 decimales
            }
        }

        private void Txt_11_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir control de borrado (Backspace)
            if (e.KeyChar == (char)Keys.Back)
            {
                return;
            }

            // Permitir solo un punto decimal
            if (e.KeyChar == '.')
            {
                if ((sender as System.Windows.Forms.TextBox).Text.Contains("."))
                {
                    e.Handled = true; // Ya hay un punto → se bloquea
                }
                return;
            }

            // Permitir solo números
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Bloquear todo lo que no sea número
            }
        }

        private void Txt_1_TextChanged(object sender, EventArgs e)
        {
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7 || cb_Area.SelectedIndex == 8)
            {
                var tb = (RadTextBox)sender;
                string original = tb.Text;
                string saneado = SanitizeNumericText(original);

                if (saneado != original)
                {
                    int sel = tb.SelectionStart;
                    tb.Text = saneado;
                    // Ajusta la posición del cursor (no exceder la longitud)
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length);
                }
                if (!string.IsNullOrEmpty(Txt_1.Text) && !string.IsNullOrEmpty(Txt_2.Text) && cb_Area.SelectedIndex == 2)
                {
                    porcentaje_aumento_humedad();
                }
                if (!string.IsNullOrEmpty(Txt_1.Text) && cb_Area.SelectedIndex == 8)
                {
                    calcular_kg_entrada_proceso();
                }
            }
            if(cb_Area.SelectedIndex == 1 && editar == true)
            {
                string valorBuscado = Txt_1.Text;

                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                if (!string.IsNullOrWhiteSpace(Txt_1.Text))
                {
                    // Consulta para buscar donde OP = valor_buscado
                    string query = "SELECT \"OP\", \"kg_frescos_enter_se\" FROM public.\"Ficha\" WHERE \"Lote\" = @valorBuscado and \"Area\" = 'Tunel/Sumergidor';";

                    // Crear parámetro
                    NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
                    };

                    // Ejecutar consulta
                    System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                    // Variable string donde guardar el resultado
                    string OP = string.Empty;
                    string kgFrescoEnterSecador = string.Empty;

                    // Verificar si se encontraron resultados
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        OP = dt.Rows[0]["OP"].ToString();
                        kgFrescoEnterSecador = dt.Rows[0]["kg_frescos_enter_se"].ToString();
                        cb_OP.Enabled = true;
                        cb_OP.Text = OP;
                        cb_OP.Enabled = false;

                        cb_Turno.Enabled = true;

                        Txt_Read_4.Text = kgFrescoEnterSecador;
                        buscar_Meta_hr();
                    }
                }
                else
                {
                    cb_Turno.SelectedIndex = -1;
                    cb_Turno.Enabled = false;
                }
            }
        }

        private void Txt_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (cb_Area.SelectedIndex == 2)
            {
                var tb = (RadTextBox)sender;
                char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

                // Permitir backspace
                if (e.KeyChar == '\b') return;

                // Permitir un solo punto decimal
                if (e.KeyChar == decimalSep)
                {
                    if (tb.Text.Contains(decimalSep)) e.Handled = true;
                    return;
                }

                // Permitir solo dígitos
                if (!char.IsDigit(e.KeyChar))
                    e.Handled = true;
            }
        }

        private void Txt_1_Validating(object sender, CancelEventArgs e)
        {
            if (cb_Area.SelectedIndex == 2 || cb_Area.SelectedIndex == 3 || cb_Area.SelectedIndex == 4 || cb_Area.SelectedIndex == 5 || cb_Area.SelectedIndex == 6 || cb_Area.SelectedIndex == 7)
            {
                RadTextBox textBox = (RadTextBox)sender;
                string text = textBox.Text;

                // Regex que acepta "" o dígitos con un punto decimal opcional
                if (!string.IsNullOrEmpty(text) && !Regex.IsMatch(text, @"^\d*\.?\d*$"))
                {
                    MetroFramework.MetroMessageBox.Show(this, "Formato inválido. Solo se permiten números con un punto decimal.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true; // mantiene el foco en el control
                }
            }
        }

        private void txt_TM_mecanico_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (MaterialTextBox)sender;
            char decimalSep = '.';

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void txt_Tiempo_energia_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (MaterialTextBox)sender;
            char decimalSep = '.';

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void dgv_metas_polvos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 6;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_polvos = dgv_metas_polvos.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_2.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_3.Text = dgv_metas_polvos.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_4.Visible = false;
                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_revolturas_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 5;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_revolturas = dgv_metas_revolturas.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_revolturas.Rows[e.RowIndex].Cells[4].Value.ToString();

                txt_Meta_5.Visible = false;

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void dgv_metas_maquinas_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                cmb_area.Enabled = true;
                cmb_area.SelectedIndex = 7;
                cmb_area.Focus();
                cmb_area.Enabled = false;

                id_global_meta_maquinas = dgv_metas_maquinas.Rows[e.RowIndex].Cells[0].Value.ToString();
                txt_op.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[1].Value.ToString();
                txt_Meta_1.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[5].Value.ToString();
                txt_Meta_2.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[2].Value.ToString();
                txt_Meta_3.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[3].Value.ToString();
                txt_Meta_4.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[4].Value.ToString();
                txt_Meta_5.Text = dgv_metas_maquinas.Rows[e.RowIndex].Cells[6].Value.ToString();

                btn_meta_edit.Enabled = true;
                btn_meta_delete.Enabled = true;
                

                txt_op.Enabled = false;
                txt_Meta_1.Enabled = false;
                txt_Meta_2.Enabled = false;
                txt_Meta_3.Enabled = false;
                txt_Meta_5.Enabled = false;
                btn_meta_save.Enabled = false;
                btn_meta_cancel.Enabled = false;
            }
        }

        private void btn_cancelar_ficha_Click(object sender, EventArgs e)
        {
            cb_Area.SelectedIndex = -1;
            reiniciarCampos();
            cb_Area.Focus();
            editar = false;
            cb_jefe_turno.Enabled = false;
            if (nivel_user == "Super Administrador")
            {
                lbl_no_emp2.Text = numero_empleado_admin;
                lbl_nom2.Text = nombre_admin;
                cb_supervisor_turno.Visible = false;

                // Cerrar el formulario Editar si está abierto
                if (_editarForm != null && !_editarForm.IsDisposed)
                {
                    _editarForm.Close();
                    _editarForm = null; // Opcional, pero recomendado para liberar la referencia
                }
            }
        }

        private void editarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            editar = true;
            borrar = false;
            // Verificar si el formulario ya existe y no está cerrado
            if (_editarForm == null || _editarForm.IsDisposed) 
            {
                // Crear nueva instancia
                _editarForm = new Editar(connectionString, editar, borrar);

                // Manejar el evento FormClosed para limpiar la referencia
                _editarForm.FormClosed += (s, args) => _editarForm = null;

                // Suscripción al evento con los dos parámetros
                _editarForm.FichaSeleccionada += (id_global, area) =>
                {

                    

                    cb_Area.SelectedIndex = -1;
                    reiniciarCampos();
                    cb_Area.Focus();
                    cb_jefe_turno.Enabled = false;
                    if (id_global == null || area == null)
                    {
                        editar = false;
                        return;
                    }
                    var materialSkinManager = MaterialSkinManager.Instance;
                    materialSkinManager.AddFormToManage(this);
                    materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
                    materialSkinManager.ColorScheme = new ColorScheme(Primary.Red600, Primary.Red600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);
                    this.Text = "Modo Edición";
                    //MessageBox.Show($"ID seleccionado: {id_global}\nÁrea: {area}");
                    id_global_ficha = id_global;
                    cb_supervisor_turno.Visible = true;
                    if (area == "Tunel/Sumergidor")
                    {
                        DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        // Consulta para buscar donde OP = valor_buscado
                        string query2 = "SELECT \"OP\", \"Turno\", \"Fecha\", \"MetaHr\", \"Hr_inicio\", \"Hr_fin\", \"Lote\", \"Kg_enter_proceso\", \"Merma_canica\", " +
                        "\"Merma_podrido\", \"Merma_tina\", \"Merma_piso\", \"Merma_canaletas\", \"Merma_lavado_bandas\", \"Personal_Operativo\", \"Cascara_carrete\", " +
                        " \"ID_Jefe\", \"ID_user\" FROM public.\"Ficha\" WHERE \"ID_Ficha\" = @valorBuscado;";

                        // Crear parámetro
                        NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id_global) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                        // Variables
                        string OP = string.Empty;
                        string Turno = string.Empty;
                        DateTime Fecha = DateTime.MinValue;
                        string Hr_inicio = string.Empty;
                        string Hr_fin = string.Empty;
                        string MetaHr = string.Empty;
                        string Lote = string.Empty;
                        string Kg_enter_proceso = string.Empty;
                        string Merma_canica = string.Empty;
                        string Merma_podrido = string.Empty;
                        string Merma_tina = string.Empty;
                        string Merma_piso = string.Empty;
                        string Merma_canaletas = string.Empty;
                        string Merma_lavado_bandas = string.Empty;
                        string Personal_Operativo = string.Empty;
                        string Cascara_carrete = string.Empty;
                        string ID_Jefe = string.Empty;
                        string ID_user = string.Empty;

                        // Verificar si se encontraron resultados
                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            OP = dt2.Rows[0]["OP"].ToString();
                            Turno = dt2.Rows[0]["Turno"].ToString();
                            Fecha = System.Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                            Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                            Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                            MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                            Lote = dt2.Rows[0]["Lote"].ToString();
                            Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                            Merma_canica = dt2.Rows[0]["Merma_canica"].ToString();
                            Merma_podrido = dt2.Rows[0]["Merma_podrido"].ToString();
                            Merma_tina = dt2.Rows[0]["Merma_tina"].ToString();
                            Merma_piso = dt2.Rows[0]["Merma_piso"].ToString();
                            Merma_canaletas = dt2.Rows[0]["Merma_canaletas"].ToString();
                            Merma_lavado_bandas = dt2.Rows[0]["Merma_lavado_bandas"].ToString();
                            Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                            Cascara_carrete = dt2.Rows[0]["Cascara_carrete"].ToString();
                            ID_Jefe = System.Convert.ToString(dt2.Rows[0]["ID_Jefe"]);
                            ID_user = System.Convert.ToString(dt2.Rows[0]["ID_user"]);
                        }
                        ////seccion para cargar supervisor
                        //buscar usuario
                        query2 = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"ID_User\" = @idUser";

                        // Crear parámetro
                        NpgsqlParameter[] parameters = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(ID_user) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt1 = dbHelper.ExecuteSelectQuery(query2, parameters);

                        string usuario_var = string.Empty;
                        string no_empleado_var = string.Empty;

                        if (dt1 != null && dt1.Rows.Count > 0)
                        {
                            usuario_var = dt1.Rows[0]["Usuario"].ToString();
                            no_empleado_var = dt1.Rows[0]["No_Empleado"].ToString();
                        }

                        lbl_no_emp2.Text = no_empleado_var;
                        lbl_nom2.Text = usuario_var;

                        string query = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"Nivel\" = 'Supervisor'";
                        dbHelper.LoadDataIntoComboBox(query, cb_supervisor_turno, "Usuario", "No_Empleado");
                        //////////////termina seccion supervisor
                        cb_Area.Text = area;
                        cb_OP.Text = OP;
                        // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                        cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                        cb_OP.Focus();

                        cb_Turno.Text = Turno;
                        cb_Turno.Focus();
                        dtp1.Value = Fecha;
                        Mask_txt_hr1.Text = Hr_inicio;
                        Mask_txt_hr2.Text = Hr_fin;
                        Txt_meta.Text = MetaHr;
                        Txt_1.Text = Lote;
                        Txt_2.Text = Kg_enter_proceso;
                        Txt_3.Text = Merma_canica;
                        Txt_4.Text = Merma_podrido;
                        Txt_5.Text = Merma_tina;
                        Txt_6.Text = Merma_piso;
                        Txt_7.Text = Merma_canaletas;
                        Txt_8.Text = Merma_lavado_bandas;
                        Txt_9.Text = Personal_Operativo;
                        Txt_10.Text = Cascara_carrete;
                        // Deshabilitar controles no editables
                        cb_Area.Enabled = false;
                        Txt_1.Enabled = false;
                        if (!string.IsNullOrWhiteSpace(ID_Jefe))
                        {
                            cb_jefe_turno.SelectedValue = System.Convert.ToInt32(ID_Jefe);
                            cb_jefe_turno.Focus();
                        }

                        actualiza_tiempos(id_global);
                    }
                    if (area == "Despegue")
                    {
                        DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        // Consulta para buscar donde OP = valor_buscado
                        string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Lote"", ""Kg_prod_seco"", ""Merma_kg"", ""Kg_fuera_espec"", ""Kg_resecar"", 
                                    ""Personal_Operativo"", ""Hr_programadas"", ""Hr_efectivas"", ""porcent_cump_meta"", ""Kg_meta"", ""Relacion_Fr_seco"", 
                                    ""FTT"", ""Hr_inicio"", ""Hr_fin"", ""MetaHr"", ""ID_Jefe"", ""ID_user"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";

                        // Crear parámetro
                        NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id_global) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);



                        // Variables
                        string OP = string.Empty;
                        string Turno = string.Empty;
                        DateTime Fecha = DateTime.MinValue;
                        string Hr_inicio = string.Empty;
                        string Hr_fin = string.Empty;
                        string MetaHr = string.Empty;
                        string Lote = string.Empty;
                        string Kg_prod_seco = string.Empty;
                        string Merma_kg = string.Empty;
                        string Kg_fuera_espec = string.Empty;
                        string Kg_resecar = string.Empty;
                        string Personal_Operativo = string.Empty;
                        string ID_Jefe = string.Empty;
                        string ID_user = string.Empty;

                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            OP = dt2.Rows[0]["OP"].ToString();
                            Turno = dt2.Rows[0]["Turno"].ToString();
                            Fecha = System.Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                            Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                            Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                            MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                            Lote = dt2.Rows[0]["Lote"].ToString();
                            Kg_prod_seco = dt2.Rows[0]["Kg_prod_seco"].ToString();
                            Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                            Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                            Kg_resecar = dt2.Rows[0]["Kg_resecar"].ToString();
                            Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                            ID_Jefe = System.Convert.ToString(dt2.Rows[0]["ID_Jefe"]);
                            ID_user = System.Convert.ToString(dt2.Rows[0]["ID_user"]);
                        }
                        ////seccion para cargar supervisor
                        //buscar usuario
                        query2 = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"ID_User\" = @idUser";

                        // Crear parámetro
                        NpgsqlParameter[] parameters = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(ID_user) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt1 = dbHelper.ExecuteSelectQuery(query2, parameters);

                        string usuario_var = string.Empty;
                        string no_empleado_var = string.Empty;

                        if (dt1 != null && dt1.Rows.Count > 0)
                        {
                            usuario_var = dt1.Rows[0]["Usuario"].ToString();
                            no_empleado_var = dt1.Rows[0]["No_Empleado"].ToString();
                        }

                        lbl_no_emp2.Text = no_empleado_var;
                        lbl_nom2.Text = usuario_var;

                        string query = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"Nivel\" = 'Supervisor'";
                        dbHelper.LoadDataIntoComboBox(query, cb_supervisor_turno, "Usuario", "No_Empleado");
                        //////////////termina seccion supervisor
                        ///
                        cb_Area.Text = area;
                        cb_OP.Text = OP;
                        // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                        cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                        cb_OP.Focus();
                        cb_Turno.Text = Turno;
                        cb_Turno.Focus();
                        dtp1.Value = Fecha;
                        Mask_txt_hr1.Text = Hr_inicio;
                        Mask_txt_hr2.Text = Hr_fin;
                        Txt_meta.Text = MetaHr;
                        Txt_1.Text = Lote;
                        Txt_2.Text = Kg_prod_seco;
                        Txt_3.Text = Merma_kg;
                        Txt_4.Text = Kg_fuera_espec;
                        Txt_5.Text = Kg_resecar;
                        Txt_6.Text = Personal_Operativo;
                        if (!string.IsNullOrWhiteSpace(ID_Jefe))
                        {
                            cb_jefe_turno.SelectedValue = System.Convert.ToInt32(ID_Jefe);
                            cb_jefe_turno.Focus();
                        }

                        // Deshabilitar controles no editables
                        cb_Area.Enabled = false;
                        radMultiColumnComboBox1.Enabled = false;

                        actualiza_tiempos(id_global);
                    }
                    if (area == "Evaporado" || area == "Grind" || area == "Empacado" || area == "Revolturas")
                    {
                        DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Kg_meta"", ""Kg_enter_proceso"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""MetaHr"", ""ID_Jefe"", ""ID_user"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                        // Crear parámetro
                        NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id_global) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);

                        // Variables
                        string OP = string.Empty;
                        string Turno = string.Empty;
                        DateTime Fecha = DateTime.MinValue;
                        string Hr_inicio = string.Empty;
                        string Hr_fin = string.Empty;
                        string MetaHr = string.Empty;
                        string Kg_meta = string.Empty;
                        string Kg_enter_proceso = string.Empty;
                        string Kg_prod_term = string.Empty;
                        string Kg_fuera_espec = string.Empty;
                        string Merma_kg = string.Empty;
                        string Personal_Operativo = string.Empty;
                        string ID_Jefe = string.Empty;
                        string ID_user = string.Empty;

                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            OP = dt2.Rows[0]["OP"].ToString();
                            Turno = dt2.Rows[0]["Turno"].ToString();
                            Fecha = System.Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                            Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                            Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                            MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                            Kg_meta = dt2.Rows[0]["Kg_meta"].ToString();
                            Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                            Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                            Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                            Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                            Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                            ID_Jefe = System.Convert.ToString(dt2.Rows[0]["ID_Jefe"]);
                            ID_user = System.Convert.ToString(dt2.Rows[0]["ID_user"]);
                        }
                        ////seccion para cargar supervisor
                        //buscar usuario
                        query2 = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"ID_User\" = @idUser";

                        // Crear parámetro
                        NpgsqlParameter[] parameters = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(ID_user) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt1 = dbHelper.ExecuteSelectQuery(query2, parameters);

                        string usuario_var = string.Empty;
                        string no_empleado_var = string.Empty;

                        if (dt1 != null && dt1.Rows.Count > 0)
                        {
                            usuario_var = dt1.Rows[0]["Usuario"].ToString();
                            no_empleado_var = dt1.Rows[0]["No_Empleado"].ToString();
                        }

                        lbl_no_emp2.Text = no_empleado_var;
                        lbl_nom2.Text = usuario_var;

                        string query = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"Nivel\" = 'Supervisor'";
                        dbHelper.LoadDataIntoComboBox(query, cb_supervisor_turno, "Usuario", "No_Empleado");
                        //////////////termina seccion supervisor
                        cb_Area.Text = area;
                        cb_OP.Text = OP;
                        // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                        cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                        cb_OP.Focus();
                        cb_Turno.Text = Turno;
                        cb_Turno.Focus();
                        dtp1.Value = Fecha;
                        Mask_txt_hr1.Text = Hr_inicio;
                        Mask_txt_hr2.Text = Hr_fin;
                        Txt_meta.Text = MetaHr;
                        Txt_1.Text = Kg_enter_proceso;
                        Txt_2.Text = Kg_prod_term;
                        Txt_3.Text = Kg_fuera_espec;
                        Txt_4.Text = Merma_kg;
                        Txt_5.Text = Personal_Operativo;
                        // Deshabilitar controles no editables
                        cb_Area.Enabled = false;
                        if (!string.IsNullOrWhiteSpace(ID_Jefe))
                        {
                            cb_jefe_turno.SelectedValue = System.Convert.ToInt32(ID_Jefe);
                            cb_jefe_turno.Focus();
                        }

                        actualiza_tiempos(id_global);
                    }
                    if (area == "Inspeccion")
                    {
                        DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Proceso"", ""Kg_prod_term"", ""Kg_enter_proceso"", 
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""MetaHr"", ""ID_Jefe"", ""ID_user"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                        // Crear parámetro
                        NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id_global) }
                        };
                        // Ejecutar consulta
                        System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                        // Variables
                        string OP = string.Empty;
                        string Turno = string.Empty;
                        DateTime Fecha = DateTime.MinValue;
                        string Hr_inicio = string.Empty;
                        string Hr_fin = string.Empty;
                        string MetaHr = string.Empty;
                        string Proceso = string.Empty;
                        string Kg_prod_term = string.Empty;
                        string Kg_enter_proceso = string.Empty;
                        string Kg_fuera_espec = string.Empty;
                        string Merma_kg = string.Empty;
                        string Personal_Operativo = string.Empty;
                        string ID_Jefe = string.Empty;
                        string ID_user = string.Empty;

                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            OP = dt2.Rows[0]["OP"].ToString();
                            Turno = dt2.Rows[0]["Turno"].ToString();
                            Fecha = System.Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                            Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                            Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                            MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                            Proceso = dt2.Rows[0]["Proceso"].ToString();
                            Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                            Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                            Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                            Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                            Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                            ID_Jefe = System.Convert.ToString(dt2.Rows[0]["ID_Jefe"]);
                            ID_user = System.Convert.ToString(dt2.Rows[0]["ID_user"]);
                        }
                        ////seccion para cargar supervisor
                        //buscar usuario
                        query2 = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"ID_User\" = @idUser";

                        // Crear parámetro
                        NpgsqlParameter[] parameters = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(ID_user) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt1 = dbHelper.ExecuteSelectQuery(query2, parameters);

                        string usuario_var = string.Empty;
                        string no_empleado_var = string.Empty;

                        if (dt1 != null && dt1.Rows.Count > 0)
                        {
                            usuario_var = dt1.Rows[0]["Usuario"].ToString();
                            no_empleado_var = dt1.Rows[0]["No_Empleado"].ToString();
                        }

                        lbl_no_emp2.Text = no_empleado_var;
                        lbl_nom2.Text = usuario_var;

                        string query = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"Nivel\" = 'Supervisor'";
                        dbHelper.LoadDataIntoComboBox(query, cb_supervisor_turno, "Usuario", "No_Empleado");
                        //////////////termina seccion supervisor
                        cb_Area.Text = area;
                        cb_OP.Text = OP;
                        // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                        cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                        cb_OP.Focus();
                        cb_Turno.Text = Turno;
                        cb_Turno.Focus();
                        dtp1.Value = Fecha;
                        Mask_txt_hr1.Text = Hr_inicio;
                        Mask_txt_hr2.Text = Hr_fin;
                        Txt_meta.Text = MetaHr;
                        Txt_1.Text = Kg_enter_proceso;
                        Txt_2.Text = Kg_prod_term;
                        Txt_3.Text = Kg_fuera_espec;
                        Txt_4.Text = Merma_kg;
                        Txt_5.Text = Personal_Operativo;
                        cb_proceso.Text = Proceso;
                        cb_proceso.Focus();
                        // Deshabilitar controles no editables
                        cb_Area.Enabled = false;
                        if (!string.IsNullOrWhiteSpace(ID_Jefe))
                        {
                            cb_jefe_turno.SelectedValue = System.Convert.ToInt32(ID_Jefe);
                            cb_jefe_turno.Focus();
                        }
                        actualiza_tiempos(id_global);
                    }
                    if (area == "Polvos")
                    {
                        DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""Kg_enter_proceso"", ""Kg_prod_term"",  
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Hr_inicio"", ""Hr_fin"", 
                    ""MetaHr"", ""Polvo_colector"", ""Granulo"", ""ID_Jefe"", ""ID_user"" FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                        // Crear parámetro
                        NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id_global) }
                        };
                        // Ejecutar consulta
                        System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                        // Variables
                        string OP = string.Empty;
                        string Turno = string.Empty;
                        DateTime Fecha = DateTime.MinValue;
                        string Hr_inicio = string.Empty;
                        string Hr_fin = string.Empty;
                        string MetaHr = string.Empty;
                        string Kg_prod_term = string.Empty;
                        string Kg_enter_proceso = string.Empty;
                        string Kg_fuera_espec = string.Empty;
                        string Merma_kg = string.Empty;
                        string Personal_Operativo = string.Empty;
                        string Polvo_colector = string.Empty;
                        string Granulo = string.Empty;
                        string ID_Jefe = string.Empty;
                        string ID_user = string.Empty;

                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            OP = dt2.Rows[0]["OP"].ToString();
                            Turno = dt2.Rows[0]["Turno"].ToString();
                            Fecha = System.Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                            Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                            Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                            MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                            Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                            Kg_enter_proceso = dt2.Rows[0]["Kg_enter_proceso"].ToString();
                            Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                            Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                            Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                            Polvo_colector = dt2.Rows[0]["Polvo_colector"].ToString();
                            Granulo = dt2.Rows[0]["Granulo"].ToString();
                            ID_Jefe = System.Convert.ToString(dt2.Rows[0]["ID_Jefe"]);
                            ID_user = System.Convert.ToString(dt2.Rows[0]["ID_user"]);
                        }
                        ////seccion para cargar supervisor
                        //buscar usuario
                        query2 = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"ID_User\" = @idUser";

                        // Crear parámetro
                        NpgsqlParameter[] parameters = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(ID_user) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt1 = dbHelper.ExecuteSelectQuery(query2, parameters);

                        string usuario_var = string.Empty;
                        string no_empleado_var = string.Empty;

                        if (dt1 != null && dt1.Rows.Count > 0)
                        {
                            usuario_var = dt1.Rows[0]["Usuario"].ToString();
                            no_empleado_var = dt1.Rows[0]["No_Empleado"].ToString();
                        }

                        lbl_no_emp2.Text = no_empleado_var;
                        lbl_nom2.Text = usuario_var;

                        string query = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"Nivel\" = 'Supervisor'";
                        dbHelper.LoadDataIntoComboBox(query, cb_supervisor_turno, "Usuario", "No_Empleado");
                        //////////////termina seccion supervisor
                        cb_Area.Text = area;
                        cb_OP.Text = OP;
                        // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                        cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                        cb_OP.Focus();
                        cb_Turno.Text = Turno;
                        cb_Turno.Focus();
                        dtp1.Value = Fecha;
                        Mask_txt_hr1.Text = Hr_inicio;
                        Mask_txt_hr2.Text = Hr_fin;
                        Txt_meta.Text = MetaHr;
                        Txt_1.Text = Kg_enter_proceso;
                        Txt_2.Text = Kg_prod_term;
                        Txt_3.Text = Kg_fuera_espec;
                        Txt_4.Text = Merma_kg;
                        Txt_5.Text = Personal_Operativo;
                        cb_proceso.Focus();
                        Txt_6.Text = Polvo_colector;
                        Txt_7.Text = Granulo;
                        // Deshabilitar controles no editables
                        cb_Area.Enabled = false;
                        if (!string.IsNullOrWhiteSpace(ID_Jefe))
                        {
                            cb_jefe_turno.SelectedValue = System.Convert.ToInt32(ID_Jefe);
                            cb_jefe_turno.Focus();
                        }
                        actualiza_tiempos(id_global);
                    }
                    if (area == "Máquinas")
                    {
                        DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                        string query2 = @"SELECT ""Fecha"", ""Turno"", ""OP"", ""MetaHr"", ""Hr_inicio"", ""Hr_fin"", ""Pz_prod"", ""Kg_meta"", ""Kg_prod_term"", 
                    ""Kg_fuera_espec"", ""Merma_kg"", ""Personal_Operativo"", ""Bobina_kg_enter"", ""Bobina_utilizada"", ""Bobina_merma"", ""ID_Jefe"", ""ID_user"" 
                    FROM public.""Ficha"" WHERE ""ID_Ficha"" = @valorBuscado;";
                        // Crear parámetro
                        NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id_global) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);

                        // Variables
                        string OP = string.Empty;
                        string Turno = string.Empty;
                        DateTime Fecha = DateTime.MinValue;
                        string Hr_inicio = string.Empty;
                        string Hr_fin = string.Empty;
                        string MetaHr = string.Empty;
                        string Kg_meta = string.Empty;
                        string Kg_prod_term = string.Empty;
                        string Kg_fuera_espec = string.Empty;
                        string Merma_kg = string.Empty;
                        string Personal_Operativo = string.Empty;
                        string Pz_prod = string.Empty;
                        string Bobina_kg_enter = string.Empty;
                        string Bobina_utilizada = string.Empty;
                        string Bobina_merma = string.Empty;
                        string ID_Jefe = string.Empty;
                        string ID_user = string.Empty;

                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            Fecha = System.Convert.ToDateTime(dt2.Rows[0]["Fecha"]);
                            Turno = dt2.Rows[0]["Turno"].ToString();
                            OP = dt2.Rows[0]["OP"].ToString();
                            MetaHr = dt2.Rows[0]["MetaHr"].ToString();
                            Hr_inicio = dt2.Rows[0]["Hr_inicio"].ToString();
                            Hr_fin = dt2.Rows[0]["Hr_fin"].ToString();
                            Pz_prod = dt2.Rows[0]["Pz_prod"].ToString();
                            Kg_meta = dt2.Rows[0]["Kg_meta"].ToString();
                            Kg_prod_term = dt2.Rows[0]["Kg_prod_term"].ToString();
                            Kg_fuera_espec = dt2.Rows[0]["Kg_fuera_espec"].ToString();
                            Merma_kg = dt2.Rows[0]["Merma_kg"].ToString();
                            Personal_Operativo = dt2.Rows[0]["Personal_Operativo"].ToString();
                            Bobina_kg_enter = dt2.Rows[0]["Bobina_kg_enter"].ToString();
                            Bobina_utilizada = dt2.Rows[0]["Bobina_utilizada"].ToString();
                            Bobina_merma = dt2.Rows[0]["Bobina_merma"].ToString();
                            ID_Jefe = System.Convert.ToString(dt2.Rows[0]["ID_Jefe"]);
                            ID_user = System.Convert.ToString(dt2.Rows[0]["ID_user"]);
                        }
                        ////seccion para cargar supervisor
                        //buscar usuario
                        query2 = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"ID_User\" = @idUser";

                        // Crear parámetro
                        NpgsqlParameter[] parameters = new NpgsqlParameter[]
                        {
                        new NpgsqlParameter("@idUser", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(ID_user) }
                        };

                        // Ejecutar consulta
                        System.Data.DataTable dt1 = dbHelper.ExecuteSelectQuery(query2, parameters);

                        string usuario_var = string.Empty;
                        string no_empleado_var = string.Empty;

                        if (dt1 != null && dt1.Rows.Count > 0)
                        {
                            usuario_var = dt1.Rows[0]["Usuario"].ToString();
                            no_empleado_var = dt1.Rows[0]["No_Empleado"].ToString();
                        }

                        lbl_no_emp2.Text = no_empleado_var;
                        lbl_nom2.Text = usuario_var;

                        string query = "SELECT \"Usuario\", \"No_Empleado\" FROM public.\"Usuarios\" where \"Nivel\" = 'Supervisor'";
                        dbHelper.LoadDataIntoComboBox(query, cb_supervisor_turno, "Usuario", "No_Empleado");
                        //////////////termina seccion supervisor
                        cb_Area.Text = area;
                        cb_OP.Text = OP;
                        // LLAMAR MANUALMENTE EL EVENTO después de asignar el valor
                        cb_OP_SelectionChangeCommitted(cb_OP, EventArgs.Empty);
                        cb_OP.Focus();
                        cb_Turno.Text = Turno;
                        cb_Turno.Focus();
                        dtp1.Value = Fecha;
                        Mask_txt_hr1.Text = Hr_inicio;
                        Mask_txt_hr2.Text = Hr_fin;
                        Txt_meta.Text = MetaHr;
                        Txt_1.Text = Pz_prod;
                        Txt_2.Text = Kg_prod_term;
                        Txt_3.Text = Kg_fuera_espec;
                        Txt_4.Text = Merma_kg;
                        Txt_5.Text = Personal_Operativo;
                        Txt_6.Text = Bobina_kg_enter;
                        Txt_7.Text = Bobina_utilizada;
                        Txt_8.Text = Bobina_merma;
                        // Deshabilitar controles no editables
                        cb_Area.Enabled = false;
                        if (!string.IsNullOrWhiteSpace(ID_Jefe))
                        {
                            cb_jefe_turno.SelectedValue = System.Convert.ToInt32(ID_Jefe);
                            cb_jefe_turno.Focus();
                        }

                        actualiza_tiempos(id_global);
                    }

                };
                _editarForm.Show();
            }
            else
            {
                // Ya existe, traer al frente
                _editarForm.BringToFront();
                _editarForm.Focus();
                var materialSkinManager = MaterialSkinManager.Instance;
                materialSkinManager.AddFormToManage(this);
                materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
                materialSkinManager.ColorScheme = new ColorScheme(Primary.Red600, Primary.Red600, Primary.BlueGrey800, Accent.Blue700, TextShade.WHITE);
                this.Text = "Modo Edición";
            }

        }

        private void actualiza_tiempos(string id)
        {
            try
            {
                // Cargar tiempos muertos mecánicos
                actualiza_tiempos_mecanicos(id);

                // Cargar tiempos muertos operativos
                actualiza_tiempos_operativos(id);

                // Cargar tiempos muertos comida (si es necesario)
                actualiza_tiempo_comida(id);

                // Cargar tiempos muertos energía (si es necesario)
                actualiza_tiempo_energia(id);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al cargar tiempos muertos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void actualiza_tiempos_mecanicos(string id)
        {
            try
            {
                dgv_mecanico.DataSource = null;
                dgv_mecanico.Rows.Clear();
                dgv_mecanico.Columns.Clear();

                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                // Consulta actualizada para incluir la columna "Tipo"
                string query = @"SELECT ""Min_Detenidos"" as ""Minutos Detenidos"", 
                        ""Motivos"" as ""Motivos Mecánicos"",
                        ""Tipo"" as ""Tipo""
                 FROM public.""Tiempo_muerto_Mecanico"" 
                 WHERE ""ID_Ficha"" = @id_ficha 
                 ORDER BY ""ID_T_Mecanico"" ASC;";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id) }
                };

                // Obtener los datos
                DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                if (dt == null || dt.Rows.Count == 0)
                {
                    // Configurar columnas vacías si no hay datos
                    ConfigurarColumnasMecanico();
                    return;
                }

                // Configurar columnas
                ConfigurarColumnasMecanico();

                // Cargar datos en el grid
                foreach (DataRow row in dt.Rows)
                {
                    int rowIndex = dgv_mecanico.Rows.Add();

                    // Minutos Detenidos
                    dgv_mecanico.Rows[rowIndex].Cells["colMinutos"].Value = row["Minutos Detenidos"];

                    // Tipo (ComboBox)
                    string tipo = row["Tipo"]?.ToString();
                    if (!string.IsNullOrEmpty(tipo))
                    {
                        dgv_mecanico.Rows[rowIndex].Cells["Column2"].Value = tipo;
                    }

                    // Motivos Mecánicos
                    dgv_mecanico.Rows[rowIndex].Cells["colMotivos"].Value = row["Motivos Mecánicos"];
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al cargar tiempos mecánicos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void actualiza_tiempos_operativos(string id)
        {
            try
            {
                dgv_operativo.DataSource = null;
                dgv_operativo.Rows.Clear();
                dgv_operativo.Columns.Clear();

                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                // Consulta actualizada para incluir la columna "Tipo"
                string query = @"SELECT ""Min_Detenidos"" as ""Minutos Detenidos"", 
                        ""Motivos"" as ""Motivos Operativos"",
                        ""Tipo"" as ""Tipo""
                 FROM public.""Tiempo_Muerto_Operativo"" 
                 WHERE ""ID_Ficha"" = @id_ficha 
                 ORDER BY ""ID_T_Operativo"" ASC;";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id) }
                };

                // Obtener los datos
                DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                if (dt == null || dt.Rows.Count == 0)
                {
                    // Configurar columnas vacías si no hay datos
                    ConfigurarColumnasOperativo();
                    return;
                }

                // Configurar columnas
                ConfigurarColumnasOperativo();

                // Cargar datos en el grid
                foreach (DataRow row in dt.Rows)
                {
                    int rowIndex = dgv_operativo.Rows.Add();

                    // Minutos Detenidos
                    dgv_operativo.Rows[rowIndex].Cells["colMinutos"].Value = row["Minutos Detenidos"];

                    // Tipo (ComboBox)
                    string tipo = row["Tipo"]?.ToString();
                    if (!string.IsNullOrEmpty(tipo))
                    {
                        dgv_operativo.Rows[rowIndex].Cells["columTipo"].Value = tipo;

                        // También configurar los motivos disponibles según el tipo
                        ConfigurarMotivosSegunTipo(rowIndex, tipo);
                    }

                    // Motivos Operativos
                    string motivo = row["Motivos Operativos"]?.ToString();
                    if (!string.IsNullOrEmpty(motivo))
                    {
                        dgv_operativo.Rows[rowIndex].Cells["columMotivo"].Value = motivo;
                    }
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al cargar tiempos operativos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConfigurarMotivosSegunTipo(int rowIndex, string tipo)
        {
            if (rowIndex < 0 || rowIndex >= dgv_operativo.Rows.Count) return;

            DataGridViewComboBoxCell cmbMotivo =
                dgv_operativo.Rows[rowIndex].Cells["columMotivo"] as DataGridViewComboBoxCell;

            if (cmbMotivo == null) return;

            cmbMotivo.Items.Clear();

            if (tipo == "ALMACÉN")
            {
                cmbMotivo.Items.Add("ESPERA DE PRODUCTO");
                cmbMotivo.Items.Add("ESPERA DE INSUMOS");
            }
            else if (tipo == "CALIDAD")
            {
                cmbMotivo.Items.Add("PRODUCTO SIN LIBERAR");
                cmbMotivo.Items.Add("ESPERA POR AW");
                cmbMotivo.Items.Add("DETENIDO POR PRODUCTO FUERA DE ESPECIFICACIÓN");
                cmbMotivo.Items.Add("PRESENCIA DE MATERIA EXTRAÑA");
            }
            else if (tipo == "PRODUCCIÓN")
            {
                cmbMotivo.Items.Add("CAMBIO DE LOTE");
                cmbMotivo.Items.Add("ORGANIZACIÓN DE ARRANQUE");
                cmbMotivo.Items.Add("ACOMODO DE PERSONAL");
                cmbMotivo.Items.Add("LIMPIEZAS");
                cmbMotivo.Items.Add("PREPARACIÓN DE ÁREA");
            }
        }

        private void ConfigurarColumnasMecanico()
        {
            // Limpiar columnas existentes
            dgv_mecanico.Columns.Clear();

            // Columna 0: Minutos Detenidos (TextBox)
            DataGridViewTextBoxColumn colMinutos = new DataGridViewTextBoxColumn();
            colMinutos.Name = "colMinutos";
            colMinutos.HeaderText = "Minutos Detenidos";
            colMinutos.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colMinutos.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv_mecanico.Columns.Add(colMinutos);

            // Columna 1: Tipo (ComboBox)
            DataGridViewComboBoxColumn colTipo = new DataGridViewComboBoxColumn();
            colTipo.Name = "Column2"; // Nombre que mencionaste
            colTipo.HeaderText = "Tipo";
            colTipo.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            colTipo.Width = 161; // Ajusta el ancho según tus necesidades

            // Agregar los items del ComboBox (mismos que configuraste)
            colTipo.Items.Add("FALLA TÉCNICA EN MAQUINARIA");
            colTipo.Items.Add("FALLA EN BANDAS");
            colTipo.Items.Add("REBABA DE METALES");
            colTipo.Items.Add("FALLA DETECTOR METALES");

            dgv_mecanico.Columns.Add(colTipo);

            // Columna 2: Motivos Mecánicos (TextBox)
            DataGridViewTextBoxColumn colMotivos = new DataGridViewTextBoxColumn();
            colMotivos.Name = "colMotivos";
            colMotivos.HeaderText = "Motivos Mecánicos";
            colMotivos.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            colMotivos.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            colMotivos.Width = 470;
            dgv_mecanico.Columns.Add(colMotivos);

            // Configurar alineación de encabezados
            dgv_mecanico.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv_mecanico.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv_mecanico.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void ConfigurarColumnasOperativo()
        {
            // Limpiar columnas existentes
            dgv_operativo.Columns.Clear();

            // Columna 0: Minutos Detenidos (TextBox)
            DataGridViewTextBoxColumn colMinutos = new DataGridViewTextBoxColumn();
            colMinutos.Name = "colMinutos";
            colMinutos.HeaderText = "Minutos Detenidos";
            colMinutos.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            colMinutos.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv_operativo.Columns.Add(colMinutos);

            // Columna 1: Tipo (ComboBox)
            DataGridViewComboBoxColumn colTipo = new DataGridViewComboBoxColumn();
            colTipo.Name = "columTipo"; // Nombre que mencionaste
            colTipo.HeaderText = "Tipo";
            colTipo.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            colTipo.Width = 161; // Ajusta el ancho según tus necesidades

            // Agregar los items del ComboBox
            colTipo.Items.Add("ALMACÉN");
            colTipo.Items.Add("CALIDAD");
            colTipo.Items.Add("PRODUCCIÓN");

            dgv_operativo.Columns.Add(colTipo);

            // Columna 2: Motivos Operativos (ComboBox)
            DataGridViewComboBoxColumn colMotivo = new DataGridViewComboBoxColumn();
            colMotivo.Name = "columMotivo"; // Nombre que mencionaste
            colMotivo.HeaderText = "Motivos Operativos";
            colMotivo.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            colMotivo.Width = 470;
            dgv_operativo.Columns.Add(colMotivo);

            // Configurar alineación de encabezados
            dgv_operativo.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv_operativo.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv_operativo.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private void actualiza_tiempo_comida(string id)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT ""Minutos_Detenidos"" 
                     FROM public.""Tiempo_Muerto_Comida"" 
                     WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id) }
            };

            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            if (dt != null && dt.Rows.Count > 0)
            {
                txt_Tiempo_comida.Text = dt.Rows[0]["Minutos_Detenidos"].ToString();
            }
            else
            {
                txt_Tiempo_comida.Text = "0";
            }
        }

        private void actualiza_tiempo_energia(string id)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            string query = @"SELECT ""Minutos_Detenidos"" 
                     FROM public.""Tiempo_Muerto_Energia"" 
                     WHERE ""ID_Ficha"" = @id_ficha;";

            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
                new NpgsqlParameter("@id_ficha", NpgsqlTypes.NpgsqlDbType.Integer) { Value = System.Convert.ToInt32(id) }
            };

            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            if (dt != null && dt.Rows.Count > 0)
            {
                txt_Tiempo_energia.Text = dt.Rows[0]["Minutos_Detenidos"].ToString();
            }
            else
            {
                txt_Tiempo_energia.Text = "0";
            }
        }
        private void dtp_calidad_ValueChanged(object sender, EventArgs e)
        {
            ActualizarNumeroSemana();
        }
        // Método para actualizar el número de semana
        private void ActualizarNumeroSemana()
        {
            try
            {
                DateTime fechaSeleccionada = dtp_polvos.Value;

                // Obtener el número de semana según el calendario gregoriano
                CultureInfo cultura = CultureInfo.CurrentCulture;
                Calendar calendario = cultura.Calendar;

                // Calcular el número de semana
                int numeroSemana = calendario.GetWeekOfYear(
                    fechaSeleccionada,
                    cultura.DateTimeFormat.CalendarWeekRule,
                    cultura.DateTimeFormat.FirstDayOfWeek
                );

                // Mostrar en el TextBox
                txt_no_semana_polvos.Text = numeroSemana.ToString();
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al calcular el número de semana: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txt_merma_calidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }
                btn_cancel_polvos.Enabled = true;
                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void btn_save_calidad_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_merma_polvos.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                DateTime fecha = dtp_polvos.Value;
                if (!string.IsNullOrEmpty(id_global_polvos_calidad))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idCalidad = System.Convert.ToInt32(id_global_polvos_calidad);

                    queryInsertUpdate = "UPDATE public.\"Limpieza_polvos\" SET \"Fecha\" = @Fecha, \"Kg_merma\" = @Kg_merma WHERE \"ID_Limpieza\" = @ID_Limpieza;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date)
                        {
                            Value = fecha
                        },
                        new NpgsqlParameter("@Kg_merma", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_merma_polvos.Text)
                        },
                        new NpgsqlParameter("@ID_Limpieza", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idCalidad  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Limpieza_polvos\" (\"Fecha\", \"Kg_merma\") VALUES (@Fecha, @Kg_merma);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                        new NpgsqlParameter("@Kg_merma", System.Convert.ToDecimal(txt_merma_polvos.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_polvos_calidad();
                    txt_merma_polvos.Text = string.Empty;
                    txt_no_semana_polvos.Text = string.Empty;
                    btn_cancel_polvos.Enabled = false;
                    btn_edit_polvos.Enabled = false;
                    id_global_polvos_calidad = string.Empty;
                    btn_save_polvos.Enabled = false;
                    txt_merma_polvos.Enabled = false;
                    dtp_polvos.Enabled = false;
                }
            }
        }

        private void dgv_metas_polvos_calidad_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DateTime Fecha = DateTime.MinValue;
            if (e.RowIndex >= 0)
            {
                id_global_polvos_calidad = dgv_polvos_calidad.Rows[e.RowIndex].Cells[0].Value.ToString();
                Fecha = System.Convert.ToDateTime(dgv_polvos_calidad.Rows[e.RowIndex].Cells[1].Value);
                txt_merma_polvos.Text = dgv_polvos_calidad.Rows[e.RowIndex].Cells[3].Value.ToString();

                dtp_polvos.Value = Fecha;
                btn_edit_polvos.Enabled = true;
                btn_delete_polvos.Enabled = true;
                txt_merma_polvos.Enabled = false;
                dtp_polvos.Enabled = false;
                btn_save_polvos.Enabled = false;
                btn_cancel_polvos.Enabled = false;
            }
        }

        private void btn_new_calidad_Click(object sender, EventArgs e)
        {
            //habilitar controles
            dtp_polvos.Enabled = true;
            txt_merma_polvos.Enabled = true;
            btn_cancel_polvos.Enabled = true;
            btn_save_polvos.Enabled = true;
            btn_edit_polvos.Enabled = false;
            btn_delete_polvos.Enabled = false;

            //limpiar campos
            txt_merma_polvos.Text = string.Empty;
            dtp_polvos.Value = DateTime.Now;

            //limpiar variables globales
            id_global_polvos_calidad = string.Empty;

            //enfocar
            txt_merma_polvos.Focus();
        }

        private void btn_edit_calidad_Click(object sender, EventArgs e)
        {
            btn_edit_polvos.Enabled = false;
            btn_save_polvos.Enabled = true;
            btn_cancel_polvos.Enabled = true;
            btn_delete_polvos.Enabled = false;
            txt_merma_polvos.Enabled = true;
            txt_merma_polvos.Focus();
        }

        private void btn_cancel_calidad_Click(object sender, EventArgs e)
        {
            btn_save_polvos.Enabled = false;
            btn_cancel_polvos.Enabled = false;
            btn_delete_polvos.Enabled = false;
            dtp_polvos.Value = DateTime.Now;
            txt_merma_polvos.Text = string.Empty;
            btn_edit_polvos.Enabled = false;
            id_global_polvos_calidad = string.Empty;
            txt_merma_polvos.Enabled = false;
            dtp_polvos.Enabled = false;
        }

        private void btn_delete_calidad_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar esta Merma de la tabla?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_calidad = System.Convert.ToInt32(id_global_polvos_calidad);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Limpieza_polvos""
                       WHERE ""ID_Limpieza"" = @id";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_calidad
                    }
                };
                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {

                    actualiza_polvos_calidad();
                    txt_merma_polvos.Text = string.Empty;
                    txt_no_semana_polvos.Text = string.Empty;
                    btn_cancel_polvos.Enabled = false;
                    btn_edit_polvos.Enabled = false;
                    id_global_polvos_calidad = string.Empty;
                    btn_save_polvos.Enabled = false;
                    btn_delete_polvos.Enabled = false;
                    txt_merma_polvos.Enabled = false;
                    dtp_polvos.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar este registro", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_new_dt_op_Click(object sender, EventArgs e)
        {
            //habilitar controles
            txt_orden_produc.Enabled = true;
            txt_producto.Enabled = true;
            txt_medida.Enabled = true;
            txt_descripcion.Enabled = true;
            txt_especificacion.Enabled = true;
            txt_ingredientes.Enabled = true;
            txt_humedad.Enabled = true;
            cb_Comercio.Enabled = true;
            cb_Manzana.Enabled = true;
            txt_analisis.Enabled = true;
            txt_area_proceso.Enabled = true;
            txt_op_origen.Enabled = true;
            txt_destino.Enabled = true;
            txt_clacificacion.Enabled = true;
            btn_cancel_dt_op.Enabled = true;
            btn_save_dt_op.Enabled = true;
            btn_edit_dt_op.Enabled = false;
            btn_delete_dt_op.Enabled = false;

            //limpiar campos
            cb_Comercio.SelectedIndex = -1;
            cb_Manzana.SelectedIndex = -1;
            txt_orden_produc.Text = string.Empty;
            txt_producto.Text = string.Empty;
            txt_medida.Text = string.Empty;
            txt_descripcion.Text = string.Empty;
            txt_especificacion.Text = string.Empty;
            txt_ingredientes.Text = string.Empty;
            txt_humedad.Text = string.Empty;
            txt_analisis.Text = string.Empty;
            txt_area_proceso.Text = string.Empty;
            txt_op_origen.Text = string.Empty;
            txt_destino.Text = string.Empty;
            txt_clacificacion.Text = string.Empty;

            //limpiar variables globales
            id_global_detalles_OP = string.Empty;

            //enfocar controles
            cb_Comercio.Focus();
            cb_Manzana.Focus();
            txt_orden_produc.Focus();
        }

        private void btn_save_dt_op_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_orden_produc.Text) || string.IsNullOrEmpty(txt_producto.Text) || string.IsNullOrEmpty(txt_medida.Text) ||
                string.IsNullOrEmpty(txt_descripcion.Text) || cb_Comercio.SelectedIndex == -1 || string.IsNullOrEmpty(txt_clacificacion.Text) || 
                string.IsNullOrEmpty(txt_destino.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                

                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                if (!string.IsNullOrEmpty(id_global_detalles_OP))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int id = System.Convert.ToInt32(id_global_detalles_OP);

                    queryInsertUpdate = "UPDATE public.\"Detalles_OP\" SET \"Orden_Produccion\" = @Orden_Produccion, \"Producto\" = @Producto, \"Medida\" = @Medida, \"Descripcion\" = @Descripcion, \"Especificacion\" = @Especificacion, \"Ingredientes\" = @Ingredientes, \"Humedad\" = @Humedad, \"Comercio\" = @Comercio, \"Manzana\" = @Manzana, \"Analisis\" = @Analisis, \"Area_Proceso\" = @Area_Proceso, \"OP_Origen\" = @OP_Origen, \"Destino\" = @Destino, \"Clasificacion\" = @Clasificacion WHERE \"ID_Dt_OP\" = @ID;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Orden_Produccion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_orden_produc.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Producto", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_producto.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Medida", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_medida.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Descripcion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_descripcion.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Especificacion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_especificacion.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Ingredientes", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_ingredientes.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Humedad", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_humedad.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Comercio", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = cb_Comercio.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Manzana", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = cb_Manzana.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Analisis", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_analisis.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Area_Proceso", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_area_proceso.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@OP_Origen", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_op_origen.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Destino", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_destino.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@Clasificacion", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = txt_clacificacion.Text.Trim().ToUpper()
                        },
                        new NpgsqlParameter("@ID", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = id  // variable convertida a int
                        }
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Verificar si ya existe
                    DatabaseHelper dbHelperop = new DatabaseHelper(connectionString);
                    string queryChecop = "SELECT COUNT(*) FROM public.\"Detalles_OP\" WHERE \"Orden_Produccion\"  ILIKE @OP;";
                    NpgsqlParameter[] parametersCheck = new NpgsqlParameter[]
                    {
                    new NpgsqlParameter("@OP", txt_orden_produc.Text.ToUpper().Trim()),
                    };
                    System.Data.DataTable dtCheckop = dbHelperop.ExecuteSelectQuery(queryChecop, parametersCheck);
                    if (dtCheckop != null && dtCheckop.Rows.Count > 0 && System.Convert.ToInt32(dtCheckop.Rows[0][0]) > 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "El OP ya existe. Por favor, elija otro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Detalles_OP\" (\"Orden_Produccion\", \"Producto\", \"Medida\", \"Descripcion\", \"Especificacion\", \"Ingredientes\", \"Humedad\", \"Comercio\", \"Manzana\", \"Analisis\", \"Area_Proceso\", \"OP_Origen\", \"Destino\", \"Clasificacion\") VALUES (@Orden_Produccion, @Producto, @Medida, @Descripcion, @Especificacion, @Ingredientes, @Humedad, @Comercio, @Manzana, @Analisis, @Area_Proceso, @OP_Origen, @Destino, @Clasificacion);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Orden_Produccion", txt_orden_produc.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Producto", txt_producto.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Medida", txt_medida.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Descripcion", txt_descripcion.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Especificacion", txt_especificacion.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Ingredientes", txt_ingredientes.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Humedad", txt_humedad.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Comercio", cb_Comercio.Text),
                        new NpgsqlParameter("@Manzana", cb_Manzana.Text),
                        new NpgsqlParameter("@Analisis", txt_analisis.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Area_Proceso", txt_area_proceso.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@OP_Origen", txt_op_origen.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Destino", txt_destino.Text.ToUpper().Trim()),
                        new NpgsqlParameter("@Clasificacion", txt_clacificacion.Text.ToUpper().Trim())
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    // Actualizar DataGridView
                    actualiza_detalles_OP();
                    limpiarCampos_detalles_OP();

                    // Deshabilitar controles después de guardar
                    btn_save_dt_op.Enabled = false;
                    btn_cancel_dt_op.Enabled = false;
                    txt_orden_produc.Enabled = false;
                    txt_producto.Enabled = false;
                    txt_medida.Enabled = false;
                    txt_descripcion.Enabled = false;
                    txt_especificacion.Enabled = false;
                    txt_ingredientes.Enabled = false;
                    txt_humedad.Enabled = false;
                    txt_analisis.Enabled = false;
                    txt_area_proceso.Enabled = false;
                    txt_op_origen.Enabled = false;
                    txt_destino.Enabled = false;
                    txt_clacificacion.Enabled = false;
                    btn_edit_dt_op.Enabled = false;
                    id_global_detalles_OP = string.Empty;

                    // Deshabilitar y enfocar comboboxes
                    cb_Comercio.Enabled = true;
                    cb_Comercio.Focus();
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = true;
                    cb_Manzana.Focus();
                    cb_Manzana.Enabled = false;
                }
            }
        }

        private void dgv_detalles_op_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (nivel_user == "Administrador" || nivel_user == "Super Administrador") 
            {
                if (e.RowIndex >= 0)
                {
                    id_global_detalles_OP = dgv_detalles_op.Rows[e.RowIndex].Cells[0].Value.ToString();
                    txt_orden_produc.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[1].Value.ToString();
                    txt_producto.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[2].Value.ToString();
                    txt_medida.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[3].Value.ToString();
                    txt_descripcion.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[4].Value.ToString();
                    txt_especificacion.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[5].Value.ToString();
                    txt_ingredientes.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[6].Value.ToString();
                    txt_humedad.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[7].Value.ToString();
                    cb_Comercio.Enabled = true;
                    cb_Comercio.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[8].Value.ToString();
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = true;
                    cb_Manzana.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[9].Value.ToString();
                    cb_Manzana.Enabled = false;
                    txt_analisis.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[10].Value.ToString();
                    txt_area_proceso.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[11].Value.ToString();
                    txt_op_origen.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[12].Value.ToString();
                    txt_destino.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[13].Value.ToString();
                    txt_clacificacion.Text = dgv_detalles_op.Rows[e.RowIndex].Cells[14].Value.ToString();

                    btn_edit_dt_op.Enabled = true;
                    btn_delete_dt_op.Enabled = true;

                    txt_orden_produc.Enabled = false;
                    txt_producto.Enabled = false;
                    txt_medida.Enabled = false;
                    txt_descripcion.Enabled = false;
                    txt_especificacion.Enabled = false;
                    txt_ingredientes.Enabled = false;
                    txt_humedad.Enabled = false;
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = false;
                    txt_analisis.Enabled = false;
                    txt_area_proceso.Enabled = false;
                    txt_op_origen.Enabled = false;
                    txt_destino.Enabled = false;
                    txt_clacificacion.Enabled = false;

                    btn_save_dt_op.Enabled = false;
                    btn_cancel_dt_op.Enabled = false;
                }
            } 
        }

        private void btn_cancel_dt_op_Click(object sender, EventArgs e)
        {
            btn_save_dt_op.Enabled = false;
            btn_cancel_dt_op.Enabled = false;
            btn_delete_dt_op.Enabled = false;
            limpiarCampos_detalles_OP();
            txt_orden_produc.Enabled = false;
            txt_producto.Enabled = false;
            txt_medida.Enabled = false;
            txt_descripcion.Enabled = false;
            txt_especificacion.Enabled = false;
            txt_ingredientes.Enabled = false;
            txt_humedad.Enabled = false;
            cb_Comercio.Enabled = false;
            cb_Manzana.Enabled = false;
            txt_analisis.Enabled = false;
            txt_area_proceso.Enabled = false;
            txt_op_origen.Enabled = false;
            txt_destino.Enabled = false;
            txt_clacificacion.Enabled = false;
            btn_edit_dt_op.Enabled = false;
            id_global_detalles_OP = string.Empty;
        }

        private void btn_edit_dt_op_Click(object sender, EventArgs e)
        {
            btn_save_dt_op.Enabled = true;
            btn_cancel_dt_op.Enabled = true;
            btn_delete_dt_op.Enabled = false;
            txt_orden_produc.Enabled = false;
            txt_producto.Enabled = true;
            txt_medida.Enabled = true;
            txt_descripcion.Enabled = true;
            txt_especificacion.Enabled = true;
            txt_ingredientes.Enabled = true;
            txt_humedad.Enabled = true;
            cb_Comercio.Enabled = true;
            cb_Manzana.Enabled = true;
            txt_analisis.Enabled = true;
            txt_area_proceso.Enabled = true;
            txt_op_origen.Enabled = true;
            txt_destino.Enabled = true;
            txt_clacificacion.Enabled = true;
            txt_orden_produc.Focus();
        }

        private void btn_delete_dt_op_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar esta Orden de Producción?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_OP = System.Convert.ToInt32(id_global_detalles_OP);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Detalles_OP""
                       WHERE ""ID_Dt_OP"" = @idOP";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@idOP", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_OP
                    }
                };
                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {
                    actualiza_detalles_OP();
                    limpiarCampos_detalles_OP();
                    btn_save_dt_op.Enabled = false;
                    btn_cancel_dt_op.Enabled = false;
                    btn_delete_dt_op.Enabled = false;
                    txt_orden_produc.Enabled = false;
                    txt_producto.Enabled = false;
                    txt_medida.Enabled = false;
                    txt_descripcion.Enabled = false;
                    txt_especificacion.Enabled = false;
                    txt_ingredientes.Enabled = false;
                    btn_edit_dt_op.Enabled = false;
                    txt_analisis.Enabled = false;
                    txt_area_proceso.Enabled = false;
                    txt_op_origen.Enabled = false;
                    txt_destino.Enabled = false;
                    txt_clacificacion.Enabled = false;
                    id_global_detalles_OP = string.Empty;

                    cb_Comercio.Enabled = true;
                    cb_Comercio.Focus();
                    cb_Comercio.Enabled = false;
                    cb_Manzana.Enabled = true;
                    cb_Manzana.Focus();
                    cb_Manzana.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar la Orden de Producción", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_buscar_detalles_Click(object sender, EventArgs e)
        {
            // Si el filtro está activo, limpiar filtro y campos de búsqueda
            if (filtrodetallesOP)
            {
                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_detalles_op.Rows)
                {
                    row.Visible = true;
                }

                // Restablecer el estado del filtro
                filtrodetallesOP = false;
            }

            string OP = txt_buscar_dt_op.Text.Trim().ToUpper();
            //validar que el campo OP este lleno
            if (string.IsNullOrEmpty(OP))
            {
                MetroFramework.MetroMessageBox.Show(this, "Favor de llenar el campo OP para realizar la búsqueda.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Deseleccionar cualquier celda y fila antes de filtrar
            dgv_detalles_op.ClearSelection();
            dgv_detalles_op.CurrentCell = null;

            // Reiniciar completamente el CurrencyManager
            if (dgv_detalles_op.DataSource != null)
            {
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_detalles_op.DataSource];
                cm.SuspendBinding();

                bool hayFiltro = false;
                foreach (DataGridViewRow row in dgv_detalles_op.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool match = true;

                    if (!string.IsNullOrEmpty(OP))
                    {
                        match &= row.Cells[1].Value != null && row.Cells[1].Value.ToString().Equals(OP, StringComparison.OrdinalIgnoreCase);
                    }

                    row.Visible = match;
                    if (match) hayFiltro = true;
                }

                cm.ResumeBinding();

                // Mover a una posición válida después del filtrado
                if (cm.Count > 0)
                    cm.Position = 0;
                else
                    cm.Position = -1;

                filtroUsuariosActivo = hayFiltro;

                if (!hayFiltro)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontró OP",
                                                        "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_limpiar_detalles_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgv_detalles_op.Rows)
            {
                row.Visible = true;
            }

            // Limpiar los campos de búsqueda
            txt_buscar_dt_op.Text = string.Empty;
            txt_buscar_dt_op.Focus();

            // Restablecer el estado del filtro
            filtrodetallesOP = false;
        }

        private void Txt_3_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void Txt_4_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void Txt_5_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (RadTextBox)sender;
            char decimalSep = '.'; // si quieres respetar cultura: CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]

            // Permitir backspace
            if (e.KeyChar == '\b') return;

            // Permitir un solo punto decimal
            if (e.KeyChar == decimalSep)
            {
                if (tb.Text.Contains(decimalSep)) e.Handled = true;
                return;
            }

            // Permitir solo dígitos
            if (!char.IsDigit(e.KeyChar))
                e.Handled = true;
        }

        private void dtp_tunel_ValueChanged(object sender, EventArgs e)
        {
            ActualizarNumeroSemana_tunel();
        }
        private void ActualizarNumeroSemana_tunel()
        {
            try
            {
                DateTime fechaSeleccionada = dtp_polvos.Value;

                // Obtener el número de semana según el calendario gregoriano
                CultureInfo cultura = CultureInfo.CurrentCulture;
                Calendar calendario = cultura.Calendar;

                // Calcular el número de semana
                int numeroSemana = calendario.GetWeekOfYear(
                    fechaSeleccionada,
                    cultura.DateTimeFormat.CalendarWeekRule,
                    cultura.DateTimeFormat.FirstDayOfWeek
                );

                // Mostrar en el TextBox
                txt_no_semana_tunel.Text = numeroSemana.ToString();
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al calcular el número de semana: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txt_merma_tunel_KeyPress(object sender, KeyPressEventArgs e)
        {
            RadTextBox textBox = (RadTextBox)sender;
            string currentText = textBox.Text;

            // Permitir teclas de control (backspace, delete, etc.)
            if (char.IsControl(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
                return;
            }

            // Permitir dígitos
            if (char.IsDigit(e.KeyChar))
            {
                btn_cancel_polvos.Enabled = true;
                return;
            }

            // Validar el punto decimal
            if (e.KeyChar == '.')
            {
                // No permitir más de un punto
                if (currentText.Contains('.'))
                {
                    e.Handled = true;
                    return;
                }

                // No permitir punto al inicio
                if (currentText.Length == 0)
                {
                    e.Handled = true;
                    return;
                }
                btn_cancel_polvos.Enabled = true;
                // Permitir el punto
                return;
            }

            // Si llegó aquí, no es un carácter válido
            e.Handled = true;
        }

        private void btn_new_tunel_Click(object sender, EventArgs e)
        {
            //habilitar controles
            dtp_tunel.Enabled = true;
            txt_merma_tunel.Enabled = true;
            btn_cancel_tunel.Enabled = true;
            btn_save_tunel.Enabled = true;
            btn_edit_tunel.Enabled = false;
            btn_delete_tunel.Enabled = false;

            //limpiar campos
            txt_merma_tunel.Text = string.Empty;
            dtp_tunel.Value = DateTime.Now;

            //limpiar variables globales
            id_global_tunel_calidad = string.Empty;

            //enfocar
            txt_merma_tunel.Focus();
        }

        private void btn_save_tunel_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt_merma_tunel.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string queryInsertUpdate = string.Empty;
                int result;
                DateTime fecha = dtp_tunel.Value;
                if (!string.IsNullOrEmpty(id_global_tunel_calidad))
                {
                    // actualizar
                    // Convertir el ID a entero ANTES de crear el parámetro
                    int idCalidad = System.Convert.ToInt32(id_global_tunel_calidad);

                    queryInsertUpdate = "UPDATE public.\"Limpieza_tunel\" SET \"Fecha\" = @Fecha, \"Kg_merma\" = @Kg_merma WHERE \"ID_Tunel\" = @ID_Limpieza;";

                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date)
                        {
                            Value = fecha
                        },
                        new NpgsqlParameter("@Kg_merma", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_merma_tunel.Text)
                        },
                        new NpgsqlParameter("@ID_Limpieza", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idCalidad  // variable convertida a int
                        }
                    };

                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }
                else
                {
                    // Insertar
                    queryInsertUpdate = "INSERT INTO public.\"Limpieza_tunel\" (\"Fecha\", \"Kg_merma\") VALUES (@Fecha, @Kg_merma);";
                    NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                    {
                        new NpgsqlParameter("@Fecha", NpgsqlTypes.NpgsqlDbType.Date) { Value = fecha },
                        new NpgsqlParameter("@Kg_merma", System.Convert.ToDecimal(txt_merma_tunel.Text))
                    };
                    result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                }

                if (result > 0)
                {
                    actualiza_tunel_calidad();
                    txt_merma_tunel.Text = string.Empty;
                    txt_no_semana_tunel.Text = string.Empty;
                    btn_cancel_tunel.Enabled = false;
                    btn_edit_tunel.Enabled = false;
                    id_global_tunel_calidad = string.Empty;
                    btn_save_tunel.Enabled = false;
                    txt_merma_tunel.Enabled = false;
                    dtp_tunel.Enabled = false;
                }
            }
        }

        private void btn_edit_tunel_Click(object sender, EventArgs e)
        {
            btn_edit_tunel.Enabled = false;
            btn_save_tunel.Enabled = true;
            btn_cancel_tunel.Enabled = true;
            btn_delete_tunel.Enabled = false;
            txt_merma_tunel.Enabled = true;
            txt_merma_tunel.Focus();
        }

        private void btn_delete_tunel_Click(object sender, EventArgs e)
        {
            if (MetroFramework.MetroMessageBox.Show(this, "Presione Yes para confimar ó Presione No para cancelar", "¿Esta realmente seguro que desea borrar esta Merma de la tabla?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int id_calidad = System.Convert.ToInt32(id_global_tunel_calidad);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                string query = @"DELETE FROM public.""Limpieza_tunel""
                       WHERE ""ID_Tunel"" = @id";

                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Integer)
                    {
                        Value = id_calidad
                    }
                };
                int result = dbHelper.ExecuteNonQuery(query, parameters);

                if (result > 0)
                {

                    actualiza_tunel_calidad();
                    txt_merma_tunel.Text = string.Empty;
                    txt_no_semana_tunel.Text = string.Empty;
                    btn_cancel_tunel.Enabled = false;
                    btn_edit_tunel.Enabled = false;
                    id_global_tunel_calidad = string.Empty;
                    btn_save_tunel.Enabled = false;
                    btn_delete_tunel.Enabled = false;
                    txt_merma_tunel.Enabled = false;
                    dtp_tunel.Enabled = false;
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se pudo eliminar este registro", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btn_cancel_tunel_Click(object sender, EventArgs e)
        {
            btn_save_tunel.Enabled = false;
            btn_cancel_tunel.Enabled = false;
            btn_delete_tunel.Enabled = false;
            dtp_tunel.Value = DateTime.Now;
            txt_merma_tunel.Text = string.Empty;
            btn_edit_tunel.Enabled = false;
            id_global_tunel_calidad = string.Empty;
            txt_merma_tunel.Enabled = false;
            dtp_tunel.Enabled = false;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(tabControl1.SelectedIndex == 0) 
            {
                tab_limpiezas.SelectedIndex = 0;
            }
            if (tabControl1.SelectedIndex == 1)
            {
                tab_limpiezas.SelectedIndex = 1;
            }
        }

        private void tab_limpiezas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tab_limpiezas.SelectedIndex == 0)
            {
                tabControl1.SelectedIndex = 0;
            }
            if (tab_limpiezas.SelectedIndex == 1)
            {
                tabControl1.SelectedIndex = 1;
            }
        }

        private void dgv_Tunel_calidad_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DateTime Fecha = DateTime.MinValue;
            if (e.RowIndex >= 0)
            {
                id_global_tunel_calidad = dgv_Tunel_calidad.Rows[e.RowIndex].Cells[0].Value.ToString();
                Fecha = System.Convert.ToDateTime(dgv_Tunel_calidad.Rows[e.RowIndex].Cells[1].Value);
                txt_merma_tunel.Text = dgv_Tunel_calidad.Rows[e.RowIndex].Cells[3].Value.ToString();

                dtp_tunel.Value = Fecha;
                btn_edit_tunel.Enabled = true;
                btn_delete_tunel.Enabled = true;
                txt_merma_tunel.Enabled = false;
                dtp_tunel.Enabled = false;
                btn_save_tunel.Enabled = false;
                btn_cancel_tunel.Enabled = false;
            }
        }

        private void btn_new_report_merma_Click(object sender, EventArgs e)
        {
            actualiza_reporte_merma();
            btn_clean_merma.Enabled = true;
            btn_export_excel_merma.Enabled = true;
        }

        private void btn_clean_merma_Click(object sender, EventArgs e)
        {
            dgv_reporte_merma.DataSource = null;   // Desvincula cualquier origen de datos
            dgv_reporte_merma.Rows.Clear();        // Limpia todas las filas (por si no hay DataSource)
            dgv_reporte_merma.Columns.Clear();     // Limpia todas las columnas
            btn_export_excel_merma.Enabled = false;
            btn_clean_merma.Enabled = false;
        }
        private void ExportarDataGridViewFiltradoAExcel_ClosedXML(
    DataGridView dgv,
    string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte");

                int colIndex = 1;

                // ==========================
                // ENCABEZADOS (COLUMNAS VISIBLES)
                // ==========================
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (dgv.Columns[i].Visible)
                    {
                        var cell = worksheet.Cell(1, colIndex);
                        cell.Value = dgv.Columns[i].HeaderText;

                        cell.Style.Fill.BackgroundColor = XLColor.Orange;
                        cell.Style.Font.Bold = true;
                        cell.Style.Font.FontColor = XLColor.Black;
                        cell.Style.Alignment.Horizontal =
                            XLAlignmentHorizontalValues.Center;

                        colIndex++;
                    }
                }

                int rowIndex = 2;
                bool hayDatos = false;

                // ==========================
                // FILAS VISIBLES (RESPETA FILTROS)
                // ==========================
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (!row.IsNewRow && row.Visible)
                    {
                        colIndex = 1;

                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (dgv.Columns[j].Visible)
                            {
                                worksheet.Cell(rowIndex, colIndex).Value =
                                    row.Cells[j].Value?.ToString();

                                colIndex++;
                            }
                        }

                        rowIndex++;
                        hayDatos = true;
                    }
                }

                // ==========================
                // MENSAJE SI NO HAY DATOS
                // ==========================
                if (!hayDatos)
                {
                    worksheet.Cell(2, 1).Value =
                        "No hay datos visibles para exportar";
                }

                // ==========================
                // BORDES Y AJUSTES
                // ==========================
                var usedRange = worksheet.Range(
                    1, 1,
                    Math.Max(rowIndex - 1, 2),
                    colIndex - 1
                );

                usedRange.Style.Border.OutsideBorder =
                    XLBorderStyleValues.Thin;
                usedRange.Style.Border.InsideBorder =
                    XLBorderStyleValues.Thin;

                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
            }
        }

        private async void btn_export_excel_merma_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarDataGridViewAExcel_ClosedXML(
                            dgv_reporte_merma,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }
        private void ExportarDataGridViewAExcel_ClosedXML(
    DataGridView dgv,
    string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte");

                int colIndex = 1;

                // Encabezados (solo columnas visibles)
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (dgv.Columns[i].Visible)
                    {
                        var cell = worksheet.Cell(1, colIndex);
                        cell.Value = dgv.Columns[i].HeaderText;

                        cell.Style.Fill.BackgroundColor = XLColor.Orange;
                        cell.Style.Font.Bold = true;
                        cell.Style.Font.FontColor = XLColor.Black;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        colIndex++;
                    }
                }

                int rowIndex = 2;

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (!row.IsNewRow && row.Visible)
                    {
                        colIndex = 1;

                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (dgv.Columns[j].Visible)
                            {
                                worksheet.Cell(rowIndex, colIndex).Value =
                                    row.Cells[j].Value?.ToString();

                                colIndex++;
                            }
                        }

                        rowIndex++;
                    }
                }

                // Bordes
                var usedRange = worksheet.Range(
                    1, 1,
                    rowIndex - 1, colIndex - 1
                );

                usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
            }
        }


        private void btn_new_report_consolidado_Click(object sender, EventArgs e)
        {
            // Validar que la fecha inicial no sea mayor que la final
            if (DTP_Consolidado_1.Value.Date > DTP_Consolidado_2.Value.Date)
            {
                MetroFramework.MetroMessageBox.Show(this,"La fecha inicial no puede ser mayor que la fecha final.",
                                "Validación de Fechas",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }

            reporte_consolidado();
            btn_export_excel_consolidado.Enabled = true;
            btn_clean_consolidado.Enabled = true;
            txt_filtro_report_consolidado.Enabled = true;
            btn_filtro_consolidado.Enabled = true;
        }

        private void reporte_consolidado()
        {
            string var1 = cb_area_reporte.Text;
            string querySimple = string.Empty;

            // Obtener las fechas de los controles MetroDateTime
            DateTime fechaInicio = DTP_Consolidado_1.Value.Date;
            DateTime fechaFin = DTP_Consolidado_2.Value.Date; // Incluye todo el día final

            rgv_reporte_consolidado.DataSource = null;
            rgv_reporte_consolidado.Rows.Clear();
            rgv_reporte_consolidado.Columns.Clear();
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Crear parámetros para las fechas
            NpgsqlParameter[] parameters = null;

            if (var1 == "Tunel/Sumergidor")
            {
                querySimple = @"
WITH turnos_trabajados AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        COUNT(*) AS total_turnos_trabajados
    FROM public.""Ficha""
    WHERE ""Area"" = 'Tunel/Sumergidor'
    AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
),
merma_semanal AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        SUM(""Kg_merma"") AS merma_total_semanal
    FROM public.""Limpieza_tunel""
    WHERE ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
)
SELECT
    f.""ID_Ficha"",
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    u.""Usuario"" AS ""Supervisor"",
    jefe.""Usuario"" AS ""Jefe de Turno"",
    f.""Lote"",
    f.""OP"",
    f.""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
    f.""kg_frescos_enter_se"" AS ""Kg Frescos Entrada a Secador"",
    f.""Merma_canica"" AS ""Canica(Kg)"",
    f.""Merma_podrido"" AS ""Merma Podrido(Kg)"",
    f.""Merma_tina"" AS ""Merma Tina(Kg)"",
    f.""Merma_piso"" AS ""Merma Piso(Kg)"",
    f.""Merma_canaletas"" AS ""Merma Canaletas(Kg)"",
    f.""Merma_lavado_bandas"" AS ""Merma Lavado Bandas(Kg)"",
    f.""Cascara_carrete"" AS ""Cáscara Carrete(Kg)"",
    f.""Personal_Operativo"" as ""Personal Operativo"",
    f.""Hr_efectivas"" as ""Horas Reales"",
    
    -- Limpieza Túnel (distribuida por turnos trabajados)
    CASE 
        WHEN tt.total_turnos_trabajados > 0 AND ms.merma_total_semanal IS NOT NULL
        THEN ROUND(ms.merma_total_semanal / tt.total_turnos_trabajados, 2)
        ELSE 0 
    END AS ""Limpieza Túnel"",
    
    -- Tiempo Muerto Operativo (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmo.total_min_operativo / 60.0, 2), 0) AS ""Tiempo Muerto Operativo(Hrs)"",
    
    -- Tiempo Muerto Mecánico (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmm.total_min_mecanico / 60.0, 2), 0) AS ""Tiempo Muerto Mecánico(Hrs)""

FROM public.""Ficha"" f
LEFT JOIN turnos_trabajados tt 
    ON EXTRACT(WEEK FROM f.""Fecha"") = tt.semana 
    AND EXTRACT(YEAR FROM f.""Fecha"") = tt.año
LEFT JOIN merma_semanal ms 
    ON EXTRACT(WEEK FROM f.""Fecha"") = ms.semana 
    AND EXTRACT(YEAR FROM f.""Fecha"") = ms.año
LEFT JOIN public.""Usuarios"" u 
    ON f.""ID_user"" = u.""ID_User""
LEFT JOIN public.""Usuarios"" jefe
    ON f.""ID_Jefe"" = jefe.""ID_User""
    
-- Subconsulta para sumar tiempo muerto operativo por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para sumar tiempo muerto mecánico por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE f.""Area"" = 'Tunel/Sumergidor'
AND f.""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY f.""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Despegue")
            {
                querySimple = @"
SELECT
    f.""ID_Ficha"",
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    u.""Usuario"" AS ""Supervisor"",
    jefe.""Usuario"" AS ""Jefe de Turno"",
    f.""Lote"",
    f.""OP"",
    f.""kg_frescos_enter_se"" AS ""Kg Frescos Entrada a Secador"",
    f.""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
    f.""Kg_prod_seco"" AS ""Kilos Producto Seco"",
    f.""Merma_kg"" AS ""Merma(Kg)"",
    f.""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
    f.""Kg_resecar"" AS ""Kg para Resecar"",
    f.""Relacion_Fr_seco"" AS ""Relación Fresco-Seco"",
    f.""Personal_Operativo"" as ""Personal Operativo"",
    f.""FTT"",
    f.""Hr_efectivas"" as ""Horas Reales"",
    
    -- Tiempo Muerto Operativo (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmo.total_min_operativo / 60.0, 2), 0) AS ""Tiempo Muerto Operativo(Hrs)"",
    
    -- Tiempo Muerto Mecánico (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmm.total_min_mecanico / 60.0, 2), 0) AS ""Tiempo Muerto Mecánico(Hrs)""

FROM public.""Ficha"" f
LEFT JOIN public.""Usuarios"" u 
    ON f.""ID_user"" = u.""ID_User""
LEFT JOIN public.""Usuarios"" jefe
    ON f.""ID_Jefe"" = jefe.""ID_User""
    
-- Subconsulta para sumar tiempo muerto operativo por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para sumar tiempo muerto mecánico por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE f.""Area"" = 'Despegue'
AND f.""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY f.""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Evaporado")
            {
                querySimple = @"
SELECT
    f.""ID_Ficha"",
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    u.""Usuario"" AS ""Supervisor"",
    jefe.""Usuario"" AS ""Jefe de Turno"",
    f.""OP"",
    f.""Kg_meta"" as ""Meta(Kg)"",
    f.""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
    f.""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
    f.""Kg_prod_term"" as ""Kg Producto Terminado"",
    f.""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
    f.""Merma_kg"" AS ""Merma(Kg)"",
    f.""porcent_aumento_hum"" as ""% Aumento de Humedad"",
    f.""Personal_Operativo"" as ""Personal Operativo"",
    f.""Hr_efectivas"" as ""Horas Reales"",
    
    -- Tiempo Muerto Operativo (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmo.total_min_operativo / 60.0, 2), 0) AS ""Tiempo Muerto Operativo(Hrs)"",
    
    -- Tiempo Muerto Mecánico (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmm.total_min_mecanico / 60.0, 2), 0) AS ""Tiempo Muerto Mecánico(Hrs)""

FROM public.""Ficha"" f
LEFT JOIN public.""Usuarios"" u 
    ON f.""ID_user"" = u.""ID_User""
LEFT JOIN public.""Usuarios"" jefe
    ON f.""ID_Jefe"" = jefe.""ID_User""
    
-- Subconsulta para sumar tiempo muerto operativo por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para sumar tiempo muerto mecánico por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE f.""Area"" = 'Evaporado'
AND f.""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY f.""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Grind" || var1 == "Inspeccion" || var1 == "Empacado" || var1 == "Revolturas")
            {
                querySimple = @"
SELECT
    f.""ID_Ficha"",
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    u.""Usuario"" AS ""Supervisor"",
    jefe.""Usuario"" AS ""Jefe de Turno"",
    f.""OP"",
    f.""Kg_meta"" as ""Meta(Kg)"",
    f.""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
    f.""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
    f.""Kg_prod_term"" as ""Kg Producto Terminado"",
    f.""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
    f.""Merma_kg"" AS ""Merma(Kg)"",
    f.""Personal_Operativo"" as ""Personal Operativo"",
    f.""Hr_efectivas"" as ""Horas Reales"",
    
    -- Totales de tiempo muerto
    COALESCE(tmo.total_min_operativo, 0) AS ""Tiempo Muerto Operativo(Hrs)"",
    COALESCE(tmm.total_min_mecanico, 0) AS ""Tiempo Muerto Mecánico(Hrs)""

FROM public.""Ficha"" f
LEFT JOIN public.""Usuarios"" u 
    ON f.""ID_user"" = u.""ID_User""
LEFT JOIN public.""Usuarios"" jefe
    ON f.""ID_Jefe"" = jefe.""ID_User""
    
-- Subconsulta para sumar tiempo muerto operativo por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        Round(SUM(""Min_Detenidos"")/60.0, 2) AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para sumar tiempo muerto mecánico por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        Round(SUM(""Min_Detenidos"")/60.0, 2) AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE f.""Area"" = @Area
AND f.""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY f.""ID_Ficha"", f.""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Máquinas")
            {
                querySimple = @"
SELECT
    f.""ID_Ficha"",
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    u.""Usuario"" AS ""Supervisor"",
    jefe.""Usuario"" AS ""Jefe de Turno"",
    f.""OP"",
    f.""Kg_meta"" as ""Meta(Kg)"",
    f.""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
    f.""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
    f.""Kg_prod_term"" as ""Kg Producto Terminado"",
    f.""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
    f.""Merma_kg"" AS ""Merma(Kg)"",
    f.""Personal_Operativo"" AS ""Personal Operativo"",
    f.""Bobina_kg_enter"" AS ""Bobina Kg Entrada"",
    f.""Bobina_utilizada"" AS ""Bobina Utilizada"",
    f.""Bobina_merma"" AS ""Bobina Merma"",
    f.""Hr_efectivas"" as ""Horas Reales"",
    
    -- Tiempo Muerto Operativo (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmo.total_min_operativo / 60.0, 2), 0) AS ""Tiempo Muerto Operativo(Hrs)"",
    
    -- Tiempo Muerto Mecánico (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmm.total_min_mecanico / 60.0, 2), 0) AS ""Tiempo Muerto Mecánico(Hrs)""

FROM public.""Ficha"" f
LEFT JOIN public.""Usuarios"" u 
    ON f.""ID_user"" = u.""ID_User""
LEFT JOIN public.""Usuarios"" jefe
    ON f.""ID_Jefe"" = jefe.""ID_User""
    
-- Subconsulta para sumar tiempo muerto operativo por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para sumar tiempo muerto mecánico por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE f.""Area"" = 'Máquinas'
AND f.""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY f.""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Polvos")
            {
                querySimple = @"
WITH turnos_trabajados AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        COUNT(*) AS total_turnos_trabajados
    FROM public.""Ficha""
    WHERE ""Area"" = 'Polvos'
    AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
),
merma_semanal AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") AS semana,
        EXTRACT(YEAR FROM ""Fecha"") AS año,
        SUM(""Kg_merma"") AS merma_total_semanal
    FROM public.""Limpieza_polvos""
    WHERE ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
    GROUP BY EXTRACT(WEEK FROM ""Fecha""), EXTRACT(YEAR FROM ""Fecha"")
)
SELECT
    f.""ID_Ficha"",
    f.""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    f.""Turno"",
    u.""Usuario"" AS ""Supervisor"",
    jefe.""Usuario"" AS ""Jefe de Turno"",
    f.""OP"",
    f.""Kg_meta"" as ""Meta(Kg)"",
    f.""porcent_cump_meta"" AS ""% Cumplimiento a Metas"",
    f.""Kg_enter_proceso"" AS ""Kg Entrada(Proceso)"",
    f.""Kg_prod_term"" as ""Kg Producto Terminado"",
    f.""Kg_fuera_espec"" AS ""Kg Fuera de Especificación"",
    f.""Merma_kg"" AS ""Merma(Kg)"",
    f.""Polvo_colector"" AS ""Polvo Colector(Kg)"",
    f.""Granulo"",
    f.""Personal_Operativo"" as ""Personal Operativo"",
    f.""Hr_efectivas"" as ""Horas Reales"",
    
    -- Limpieza Polvos (distribuida por turnos trabajados)
    CASE 
        WHEN tt.total_turnos_trabajados > 0 AND ms.merma_total_semanal IS NOT NULL
        THEN ROUND(ms.merma_total_semanal / tt.total_turnos_trabajados, 2)
        ELSE 0 
    END AS ""Limpieza Polvos"",
    
    -- Tiempo Muerto Operativo (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmo.total_min_operativo / 60.0, 2), 0) AS ""Tiempo Muerto Operativo(Hrs)"",
    
    -- Tiempo Muerto Mecánico (suma por ID_Ficha, convertido a horas)
    COALESCE(ROUND(tmm.total_min_mecanico / 60.0, 2), 0) AS ""Tiempo Muerto Mecánico(Hrs)""

FROM public.""Ficha"" f
LEFT JOIN turnos_trabajados tt 
    ON EXTRACT(WEEK FROM f.""Fecha"") = tt.semana 
    AND EXTRACT(YEAR FROM f.""Fecha"") = tt.año
LEFT JOIN merma_semanal ms 
    ON EXTRACT(WEEK FROM f.""Fecha"") = ms.semana 
    AND EXTRACT(YEAR FROM f.""Fecha"") = ms.año
LEFT JOIN public.""Usuarios"" u 
    ON f.""ID_user"" = u.""ID_User""
LEFT JOIN public.""Usuarios"" jefe
    ON f.""ID_Jefe"" = jefe.""ID_User""
    
-- Subconsulta para sumar tiempo muerto operativo por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para sumar tiempo muerto mecánico por ID_Ficha
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE f.""Area"" = 'Polvos'
AND f.""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY f.""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }

            

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, rgv_reporte_consolidado, parameters);

            // Verificar si hay datos antes de configurar columnas
            if (rgv_reporte_consolidado.Columns.Count > 0)
            {
                // Configurar el DataGridView
                rgv_reporte_consolidado.Columns[0].IsVisible = false; // Ocultar la columna ID
                rgv_reporte_consolidado.Columns["Fecha"].FormatString = "{0:dd/MM/yyyy}";

                rgv_reporte_consolidado.Columns["Fecha"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["No. Semana"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["Turno"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["OP"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["Supervisor"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["Jefe de Turno"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["Personal Operativo"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["Tiempo Muerto Operativo(Hrs)"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_consolidado.Columns["Tiempo Muerto Mecánico(Hrs)"].TextAlignment = ContentAlignment.MiddleCenter;

                if (var1 == "Despegue")
                {
                    rgv_reporte_consolidado.Columns["Kg Frescos Entrada a Secador"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Lote"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Kilos Producto Seco"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Merma(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Kg Fuera de Especificación"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Kg para Resecar"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Relación Fresco-Seco"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["FTT"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (var1 == "Tunel/Sumergidor")
                {
                    rgv_reporte_consolidado.Columns["Kg Entrada(Proceso)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Kg Frescos Entrada a Secador"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Lote"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Limpieza Túnel"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Canica(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Merma Podrido(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Merma Tina(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Merma Piso(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Merma Canaletas(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Merma Lavado Bandas(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Cáscara Carrete(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (var1 == "Máquinas")
                {
                    rgv_reporte_consolidado.Columns["Bobina Kg Entrada"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Bobina Utilizada"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Bobina Merma"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (var1 == "Polvos")
                {
                    rgv_reporte_consolidado.Columns["Polvo Colector(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Granulo"].TextAlignment = ContentAlignment.MiddleCenter;
                    rgv_reporte_consolidado.Columns["Limpieza Polvos"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                // 🔹 Formato de porcentaje para las columnas de tipo decimal
                if (rgv_reporte_consolidado.Columns.Contains("% Cumplimiento a Metas"))
                {
                    var colMeta = rgv_reporte_consolidado.Columns["% Cumplimiento a Metas"];
                    colMeta.DataType = typeof(decimal);
                    colMeta.FormatString = "{0:P0}"; // ejemplo: 0.85 → 85%
                    colMeta.TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("FTT"))
                {
                    var colFTT = rgv_reporte_consolidado.Columns["FTT"];
                    colFTT.DataType = typeof(decimal);
                    colFTT.FormatString = "{0:P0}";
                    colFTT.TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("% Aumento de Humedad"))
                {
                    var colFTT = rgv_reporte_consolidado.Columns["% Aumento de Humedad"];
                    colFTT.DataType = typeof(decimal);
                    colFTT.FormatString = "{0:P0}";
                    colFTT.TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("Meta(Kg)"))
                {
                    rgv_reporte_consolidado.Columns["Meta(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("Kg Entrada(Proceso)"))
                {
                    rgv_reporte_consolidado.Columns["Kg Entrada(Proceso)"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("Kg Producto Terminado"))
                {
                    rgv_reporte_consolidado.Columns["Kg Producto Terminado"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("Kg Fuera de Especificación"))
                {
                    rgv_reporte_consolidado.Columns["Kg Fuera de Especificación"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                if (rgv_reporte_consolidado.Columns.Contains("Merma(Kg)"))
                {
                    rgv_reporte_consolidado.Columns["Merma(Kg)"].TextAlignment = ContentAlignment.MiddleCenter;
                }

                rgv_reporte_consolidado.BestFitColumns(BestFitColumnMode.DisplayedCells);
            }

            // Limpiar el filtro al cargar nuevos datos
            txt_filtro_report_consolidado.Clear();
            rgv_reporte_costo.ClearSelection();
        }

        private async void btn_export_excel_consolidado_Click(object sender, EventArgs e)
        {
          
            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarRadGridViewFiltradoAExcel_ClosedXML(
                            rgv_reporte_consolidado,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }
        private void ExportarRadGridViewFiltradoAExcel_ClosedXML(
    RadGridView radGridView,
    string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Consolidado");

                int colIndex = 1;

                // Encabezados
                foreach (GridViewDataColumn column in radGridView.Columns)
                {
                    if (column.IsVisible)
                    {
                        var cell = worksheet.Cell(1, colIndex);
                        cell.Value = column.HeaderText;

                        cell.Style.Fill.BackgroundColor = XLColor.Orange;
                        cell.Style.Font.Bold = true;
                        cell.Style.Font.FontColor = XLColor.Black;
                        cell.Style.Alignment.Horizontal =
                            XLAlignmentHorizontalValues.Center;

                        colIndex++;
                    }
                }

                int rowIndex = 2;
                bool hayDatos = false;

                foreach (GridViewRowInfo row in radGridView.ChildRows)
                {
                    if (row.IsVisible && !(row is GridViewGroupRowInfo))
                    {
                        colIndex = 1;

                        foreach (GridViewDataColumn column in radGridView.Columns)
                        {
                            if (column.IsVisible)
                            {
                                worksheet.Cell(rowIndex, colIndex).Value =
                                    row.Cells[column.Name].Value?.ToString();

                                colIndex++;
                            }
                        }

                        rowIndex++;
                        hayDatos = true;
                    }
                }

                if (!hayDatos)
                {
                    worksheet.Cell(2, 1).Value =
                        "No hay datos visibles para exportar";
                }

                var usedRange = worksheet.Range(
                    1, 1,
                    Math.Max(rowIndex - 1, 2),
                    colIndex - 1
                );

                usedRange.Style.Border.OutsideBorder =
                    XLBorderStyleValues.Thin;
                usedRange.Style.Border.InsideBorder =
                    XLBorderStyleValues.Thin;

                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
            }
        }

        private void btn_filtro_consolidado_Click(object sender, EventArgs e)
        {
            txt_filtro_report_consolidado.Clear();
        }

        private void rgv_reporte_consolidado_CustomFiltering(object sender, GridViewCustomFilteringEventArgs e)
        {
            string textoFiltro = txt_filtro_report_consolidado.Text.Trim();

            // Si no hay texto de filtro, mostrar todas las filas
            if (string.IsNullOrEmpty(textoFiltro))
            {
                e.Visible = true;
                ResetearEstiloCeldas(e);
                return;
            }

            // Iniciar actualización para mejor rendimiento
            rgv_reporte_consolidado.BeginUpdate();

            // Por defecto ocultar la fila
            e.Visible = false;

            // Buscar en todas las celdas de la fila
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                GridViewCellInfo celda = e.Row.Cells[i];

                // Verificar si la celda tiene valor
                if (celda.Value != null)
                {
                    string textoCelda = celda.Value.ToString();

                    // Buscar coincidencia (case-insensitive)
                    if (textoCelda.IndexOf(textoFiltro, 0, StringComparison.InvariantCultureIgnoreCase) >= 0)
                    {
                        e.Visible = true; // Mostrar la fila si hay coincidencia

                        // Resaltar la celda que coincide
                        celda.Style.CustomizeFill = true;
                        celda.Style.DrawFill = true;
                        celda.Style.BackColor = Color.FromArgb(201, 252, 254); // Color azul claro
                    }
                    else
                    {
                        // Resetear estilo si no coincide
                        celda.Style.Reset();
                    }
                }
                else
                {
                    // Resetear estilo si la celda es nula
                    celda.Style.Reset();
                }
            }

            rgv_reporte_consolidado.EndUpdate(false);
        }
       
        // Método auxiliar para resetear el estilo de todas las celdas
        private void ResetearEstiloCeldas(GridViewCustomFilteringEventArgs e)
        {
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                e.Row.Cells[i].Style.Reset();
            }
        }

        private void txt_filtro_report_consolidado_TextChanged(object sender, EventArgs e)
        {
            // Refrescar el grid para aplicar el filtro
            if (rgv_reporte_consolidado != null)
            {
                rgv_reporte_consolidado.MasterTemplate.Refresh();
            }
        }

        private void btn_clean_consolidado_Click(object sender, EventArgs e)
        {
            // Limpiar el DataGridView
            rgv_reporte_consolidado.DataSource = null;
            rgv_reporte_consolidado.Rows.Clear();
            rgv_reporte_consolidado.Columns.Clear();
            // Limpiar el filtro al cargar nuevos datos
            txt_filtro_report_consolidado.Clear();
            // Deshabilitar botones hasta que se seleccione un reporte y área
            btn_export_excel_consolidado.Enabled = false;
            btn_clean_consolidado.Enabled = false;
            txt_filtro_report_consolidado.Enabled = false;
            btn_filtro_consolidado.Enabled = false;
            btn_new_report_consolidado.Enabled = false;
            DTP_Consolidado_1.Enabled = false;
            DTP_Consolidado_2.Enabled = false;
            cb_area_reporte.SelectedIndex = -1;
        }

        private void cb_area_reporte_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cb_area_reporte.SelectedIndex != -1)
            {
                btn_new_report_consolidado.Enabled = true;
                DTP_Consolidado_1.Enabled = true;
                DTP_Consolidado_2.Enabled = true;
            }
        }

        private void rgv_reporte_consolidado_CellDoubleClick(object sender, GridViewCellEventArgs e)
        {
            string id_ficha_reporte = string.Empty;
            if (e.RowIndex >= 0)
            {
                id_ficha_reporte = rgv_reporte_consolidado.Rows[e.RowIndex].Cells[0].Value.ToString();
                Tiempo_Muerto formTM = new Tiempo_Muerto(connectionString, id_ficha_reporte);
                formTM.Show(); // Muestra el formulario principal
            }
        }

        private void btn_new_concentrado_Click(object sender, EventArgs e)
        {
            actualiza_reporte_concentrado();
            btn_clean_concentrado.Enabled = true;
            btn_export_excel_concentrado.Enabled = true;
            materialCard12.Enabled = true;
        }
        private void actualiza_reporte_concentrado()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Obtener el año seleccionado del ComboBox
            string añoSeleccionado = CB_Anio_Con_Des.Text;

            // Validar que se haya seleccionado un año
            if (string.IsNullOrEmpty(añoSeleccionado))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Por favor, seleccione un año antes de generar el reporte.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // Construir la consulta con filtro de año usando parámetros
            string querySimple = @"
SELECT 
    COALESCE(q1.""Año"", q2.""Año"", q3.""Año"") AS ""Año"",
    COALESCE(q1.""No. Semana"", q2.""No. Semana"", q3.""No. Semana"") AS ""No. Semana"",
    COALESCE(q1.""Mes"", q2.""Mes"", q3.""Mes"") AS ""Mes"",
    COALESCE(q1.""OP"", q2.""OP"", q3.""OP"") AS ""OP"",
    COALESCE(q2.""Horas Programadas"", 0) AS ""Horas Programadas"",
    COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0) AS ""Horas Reales"",
    COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0) AS ""Suma de Tiempo Muerto Mecanico"",
    COALESCE(q1.""Suma de Tiempo Muerto Operativo"", 0) AS ""Suma de Tiempo Muerto Operativo"",

    CASE 
        WHEN COALESCE(q2.""Horas Programadas"", 0) > 0 THEN
            ROUND(
                100 - (
                    (COALESCE(q1.""Suma de Tiempo Muerto Operativo"", 0) +
                     COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0))
                     / COALESCE(q2.""Horas Programadas"", 0) * 100
                )::numeric
            , 2)
        ELSE 0
    END AS ""%Cumplimiento Tiempo Efectivo"",

    CASE 
        WHEN (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0)) > 0 THEN
            ROUND(COALESCE(d.""Kg_fresco_hr"", 0) * (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0)), 2)
        ELSE 0
    END AS ""Kg Fresco Meta / Hras Reales"",

    COALESCE(q3.""Kg Fresco Real"", 0) AS ""Kg Fresco Real"",

    CASE 
        WHEN (COALESCE(d.""Kg_fresco_hr"", 0) * (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0))) > 0 THEN
            LEAST(
                ROUND(
                    (COALESCE(q3.""Kg Fresco Real"", 0) /
                    (COALESCE(d.""Kg_fresco_hr"", 0) * (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0))))::numeric
                    * 100
                , 2),
            100)
        ELSE 0
    END AS ""%Cumplimiento Fresco"",
	
    CASE 
        WHEN (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0)) > 0 THEN
            ROUND(COALESCE(d.""Kg_seco_hr"", 0) * (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0)), 2)
        ELSE 0
    END AS ""Kg Seco Meta / Hras Reales"",
	
    COALESCE(q3.""Kg Seco Real"", 0) AS ""Kg Seco Real"",
	
    CASE 
        WHEN (COALESCE(d.""Kg_seco_hr"", 0) * (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0))) > 0 THEN
            LEAST(
                ROUND(
                    (COALESCE(q3.""Kg Seco Real"", 0) /
                    (COALESCE(d.""Kg_seco_hr"", 0) * (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0))))::numeric
                    * 100
                , 2),
            100)
        ELSE 0
    END AS ""%Cumplimiento Secos"",
	
    COALESCE(d.""Relacion_fr_seco"", 0) AS ""Relación Fresco-Seco Meta"",
	
    CASE 
        WHEN (COALESCE(q3.""Kg Seco Real"", 0)) > 0 THEN
            ROUND(
                (COALESCE(q3.""Kg Fresco Real"", 0) /
                (COALESCE(q3.""Kg Seco Real"", 0)))::numeric, 2)
        ELSE 0
    END AS ""Relación Fresco-Seco Real"",
	
    CASE 
        WHEN (COALESCE(q3.""Kg Seco Real"", 0)) > 0 THEN
            LEAST(
                ROUND(
                    (((COALESCE(q3.""Kg Fresco Real"", 0) /
                    (COALESCE(q3.""Kg Seco Real"", 0)))) /
                    (COALESCE(d.""Relacion_fr_seco"", 0)))::numeric
                    * 100
                , 2),
            100)
        ELSE 0
    END AS ""%Cumplimiento Relación Fresco-Seco"",
	
    COALESCE(q3.""Kg Fuera de Especificación"", 0) AS ""Kg Fuera de Especificación"",

    CASE 
        WHEN COALESCE(q3.""Kg Seco Real"", 0) > 0 THEN
            ROUND(
                ((COALESCE(q3.""Kg Seco Real"", 0) - 
                  COALESCE(q3.""Kg Fuera de Especificación"", 0))
                  / COALESCE(q3.""Kg Seco Real"", 0))::numeric * 100
            , 2)
        ELSE 0
    END AS ""FTT"",

    COALESCE(q3.""Kg para Resecar"", 0) AS ""Kg para Resecar"",

    CASE 
        WHEN COALESCE(q3.""Kg Seco Real"", 0) > 0 THEN
            LEAST(
                ROUND(
                    ((COALESCE(q3.""Kg Seco Real"", 0) - COALESCE(q3.""Kg para Resecar"", 0)) / COALESCE(q3.""Kg Seco Real"", 0))::numeric * 100
                , 2),
            100)
        ELSE 0
    END AS ""%Cumplimiento Resecado"",

    CASE 
        WHEN COALESCE(q3.""Personal Operativo Suma Despegue"", 0) = 0 AND COALESCE(q4.""Personal Operativo Suma Tunel"", 0) = 0 THEN 0
        ELSE ROUND(
            (COALESCE(q3.""Personal Operativo Suma Despegue"", 0) + COALESCE(q4.""Personal Operativo Suma Tunel"", 0))::numeric /
            NULLIF(COALESCE(q3.""Cantidad Registros Despegue"", 0) + COALESCE(q4.""Cantidad Registros Tunel"", 0), 0)
        )::integer
    END AS ""Personal Operativo Promedio"",

    ROUND(
        CASE 
            WHEN COALESCE(q3.""Personal Operativo Suma Despegue"", 0) = 0 AND COALESCE(q4.""Personal Operativo Suma Tunel"", 0) = 0 THEN 0
            ELSE (COALESCE(q3.""Personal Operativo Suma Despegue"", 0) + COALESCE(q4.""Personal Operativo Suma Tunel"", 0))::numeric /
                  NULLIF(COALESCE(q3.""Cantidad Registros Despegue"", 0) + COALESCE(q4.""Cantidad Registros Tunel"", 0), 0) *
                  (COALESCE(q2.""Horas Programadas"", 0) - COALESCE(q1.""Suma de Tiempo Muerto Mecanico"", 0))
        END,
    2) AS ""Horas Hombre"",

    COALESCE(q5.""Kg Merma en Fresco"", 0) AS ""Kg Merma en Fresco""

FROM (
    SELECT 
        EXTRACT(YEAR FROM f.""Fecha"") AS ""Año"",
        EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
        EXTRACT(MONTH FROM f.""Fecha"") AS ""MesNum"",
        CASE EXTRACT(MONTH FROM f.""Fecha"")
            WHEN 1 THEN 'Enero'
            WHEN 2 THEN 'Febrero'
            WHEN 3 THEN 'Marzo'
            WHEN 4 THEN 'Abril'
            WHEN 5 THEN 'Mayo'
            WHEN 6 THEN 'Junio'
            WHEN 7 THEN 'Julio'
            WHEN 8 THEN 'Agosto'
            WHEN 9 THEN 'Septiembre'
            WHEN 10 THEN 'Octubre'
            WHEN 11 THEN 'Noviembre'
            WHEN 12 THEN 'Diciembre'
        END AS ""Mes"",
        f.""OP"",
        ROUND(COALESCE(SUM(tmo.""Min_Detenidos"")/60.0, 0)::numeric, 2) AS ""Suma de Tiempo Muerto Operativo"",
        ROUND(COALESCE(SUM(tmm.""Min_Detenidos"")/60.0, 0)::numeric, 2) AS ""Suma de Tiempo Muerto Mecanico""
    FROM public.""Ficha"" f
    LEFT JOIN public.""Tiempo_Muerto_Operativo"" tmo 
        ON f.""ID_Ficha"" = tmo.""ID_Ficha""
    LEFT JOIN public.""Tiempo_muerto_Mecanico"" tmm 
        ON f.""ID_Ficha"" = tmm.""ID_Ficha""
    WHERE f.""Area"" IN ('Despegue')
        AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio
    GROUP BY 
        EXTRACT(YEAR FROM f.""Fecha""),
        EXTRACT(WEEK FROM f.""Fecha""),
        EXTRACT(MONTH FROM f.""Fecha""),
        f.""OP""
) q1

FULL JOIN (
    SELECT 
        EXTRACT(YEAR FROM ""Fecha"") AS ""Año"",
        EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
        EXTRACT(MONTH FROM ""Fecha"") AS ""MesNum"",
        CASE EXTRACT(MONTH FROM ""Fecha"")
            WHEN 1 THEN 'Enero'
            WHEN 2 THEN 'Febrero'
            WHEN 3 THEN 'Marzo'
            WHEN 4 THEN 'Abril'
            WHEN 5 THEN 'Mayo'
            WHEN 6 THEN 'Junio'
            WHEN 7 THEN 'Julio'
            WHEN 8 THEN 'Agosto'
            WHEN 9 THEN 'Septiembre'
            WHEN 10 THEN 'Octubre'
            WHEN 11 THEN 'Noviembre'
            WHEN 12 THEN 'Diciembre'
        END AS ""Mes"",
        ""OP"",
        SUM(""Hr_programadas"") AS ""Horas Programadas""
    FROM public.""Ficha""
    WHERE ""Area"" = 'Tunel/Sumergidor'
        AND EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY 
        EXTRACT(YEAR FROM ""Fecha""),
        EXTRACT(WEEK FROM ""Fecha""),
        EXTRACT(MONTH FROM ""Fecha""),
        ""OP""
) q2
ON q1.""Año"" = q2.""Año""
AND q1.""No. Semana"" = q2.""No. Semana""
AND q1.""Mes"" = q2.""Mes""
AND q1.""OP"" = q2.""OP""

FULL JOIN (
    SELECT 
        EXTRACT(YEAR FROM ""Fecha"") AS ""Año"",
        EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
        EXTRACT(MONTH FROM ""Fecha"") AS ""MesNum"",
        CASE EXTRACT(MONTH FROM ""Fecha"")
            WHEN 1 THEN 'Enero'
            WHEN 2 THEN 'Febrero'
            WHEN 3 THEN 'Marzo'
            WHEN 4 THEN 'Abril'
            WHEN 5 THEN 'Mayo'
            WHEN 6 THEN 'Junio'
            WHEN 7 THEN 'Julio'
            WHEN 8 THEN 'Agosto'
            WHEN 9 THEN 'Septiembre'
            WHEN 10 THEN 'Octubre'
            WHEN 11 THEN 'Noviembre'
            WHEN 12 THEN 'Diciembre'
        END AS ""Mes"",
        ""OP"",
        SUM(""kg_frescos_enter_se"") AS ""Kg Fresco Real"",
        SUM(""Kg_prod_seco"") AS ""Kg Seco Real"",
        SUM(""Kg_fuera_espec"") AS ""Kg Fuera de Especificación"",
        SUM(""Kg_resecar"") AS ""Kg para Resecar"",
        SUM(""Personal_Operativo"") AS ""Personal Operativo Suma Despegue"",
        COUNT(""Personal_Operativo"") AS ""Cantidad Registros Despegue"",
        SUM(""Merma_kg"") AS ""Kg Merma en Despegue""
    FROM public.""Ficha""
    WHERE ""Area"" = 'Despegue'
        AND EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY 
        EXTRACT(YEAR FROM ""Fecha""),
        EXTRACT(WEEK FROM ""Fecha""),
        EXTRACT(MONTH FROM ""Fecha""),
        ""OP""
) q3
ON COALESCE(q1.""Año"", q2.""Año"") = q3.""Año""
AND COALESCE(q1.""No. Semana"", q2.""No. Semana"") = q3.""No. Semana""
AND COALESCE(q1.""Mes"", q2.""Mes"") = q3.""Mes""
AND COALESCE(q1.""OP"", q2.""OP"") = q3.""OP""

LEFT JOIN (
    SELECT 
        EXTRACT(YEAR FROM ""Fecha"") AS ""Año"",
        EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
        EXTRACT(MONTH FROM ""Fecha"") AS ""MesNum"",
        CASE EXTRACT(MONTH FROM ""Fecha"")
            WHEN 1 THEN 'Enero'
            WHEN 2 THEN 'Febrero'
            WHEN 3 THEN 'Marzo'
            WHEN 4 THEN 'Abril'
            WHEN 5 THEN 'Mayo'
            WHEN 6 THEN 'Junio'
            WHEN 7 THEN 'Julio'
            WHEN 8 THEN 'Agosto'
            WHEN 9 THEN 'Septiembre'
            WHEN 10 THEN 'Octubre'
            WHEN 11 THEN 'Noviembre'
            WHEN 12 THEN 'Diciembre'
        END AS ""Mes"",
        ""OP"",
        SUM(""Personal_Operativo"") AS ""Personal Operativo Suma Tunel"",
        COUNT(""Personal_Operativo"") AS ""Cantidad Registros Tunel""
    FROM public.""Ficha""
    WHERE ""Area"" = 'Tunel/Sumergidor'
        AND EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY 
        EXTRACT(YEAR FROM ""Fecha""),
        EXTRACT(WEEK FROM ""Fecha""),
        EXTRACT(MONTH FROM ""Fecha""),
        ""OP""
) q4
ON COALESCE(q1.""Año"", q2.""Año"", q3.""Año"") = q4.""Año""
AND COALESCE(q1.""No. Semana"", q2.""No. Semana"", q3.""No. Semana"") = q4.""No. Semana""
AND COALESCE(q1.""Mes"", q2.""Mes"", q3.""Mes"") = q4.""Mes""
AND COALESCE(q1.""OP"", q2.""OP"", q3.""OP"") = q4.""OP""

LEFT JOIN (
    SELECT 
        EXTRACT(YEAR FROM ""Fecha"") AS ""Año"",
        EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
        EXTRACT(MONTH FROM ""Fecha"") AS ""MesNum"",
        CASE EXTRACT(MONTH FROM ""Fecha"")
            WHEN 1 THEN 'Enero'
            WHEN 2 THEN 'Febrero'
            WHEN 3 THEN 'Marzo'
            WHEN 4 THEN 'Abril'
            WHEN 5 THEN 'Mayo'
            WHEN 6 THEN 'Junio'
            WHEN 7 THEN 'Julio'
            WHEN 8 THEN 'Agosto'
            WHEN 9 THEN 'Septiembre'
            WHEN 10 THEN 'Octubre'
            WHEN 11 THEN 'Noviembre'
            WHEN 12 THEN 'Diciembre'
        END AS ""Mes"",
        ""OP"",
        COALESCE(SUM(""Merma_podrido""), 0) + 
        COALESCE(SUM(""Merma_tina""), 0) + 
        COALESCE(SUM(""Merma_piso""), 0) + 
        COALESCE(SUM(""Merma_canaletas""), 0) + 
        COALESCE(SUM(""Merma_lavado_bandas""), 0) AS ""Kg Merma en Fresco""
    FROM public.""Ficha""
    WHERE ""Area"" = 'Tunel/Sumergidor'
        AND EXTRACT(YEAR FROM ""Fecha"") = @Anio
    GROUP BY 
        EXTRACT(YEAR FROM ""Fecha""),
        EXTRACT(WEEK FROM ""Fecha""),
        EXTRACT(MONTH FROM ""Fecha""),
        ""OP""
) q5
ON COALESCE(q1.""Año"", q2.""Año"", q3.""Año"", q4.""Año"") = q5.""Año""
AND COALESCE(q1.""No. Semana"", q2.""No. Semana"", q3.""No. Semana"", q4.""No. Semana"") = q5.""No. Semana""
AND COALESCE(q1.""Mes"", q2.""Mes"", q3.""Mes"", q4.""Mes"") = q5.""Mes""
AND COALESCE(q1.""OP"", q2.""OP"", q3.""OP"", q4.""OP"") = q5.""OP""

LEFT JOIN public.""Deshidratado"" d
ON COALESCE(q1.""OP"", q2.""OP"", q3.""OP"", q4.""OP"", q5.""OP"") = d.""OP""

ORDER BY ""Año"", ""No. Semana"", ""OP"";";

            // Crear el parámetro para el año
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@Anio", NpgsqlTypes.NpgsqlDbType.Integer)
        {
            Value = Convert.ToInt32(añoSeleccionado)
        }
            };

            // Cargar los datos con el filtro de año
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_reporte_concentrado, parameters);

            // Quitar selección inicial
            dgv_reporte_concentrado.ClearSelection();
        }
        private async void btn_export_excel_concentrado_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarDataGridViewFiltradoAExcel_ClosedXML(
                            dgv_reporte_concentrado,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }

        private void btn_clean_concentrado_Click(object sender, EventArgs e)
        {
            dgv_reporte_concentrado.DataSource = null;   // Desvincula cualquier origen de datos
            dgv_reporte_concentrado.Rows.Clear();        // Limpia todas las filas (por si no hay DataSource)
            dgv_reporte_concentrado.Columns.Clear();     // Limpia todas las columnas
            btn_export_excel_concentrado.Enabled = false;
            btn_clean_concentrado.Enabled = false;
            // Limpiar el campo de texto de búsqueda
            txt_filtro_report_consentrado.Text = "";

            // Limpiar la selección del ComboBox de semana
            cb_semana_concentrado_d.SelectedIndex = -1;
            cb_semana_concentrado_d.Focus();
            materialCard12.Enabled = false;
        }

        private void btn_buscar_Concen_Des_Click(object sender, EventArgs e)
        {
            // Si el filtro está activo, limpiar filtro y campos de búsqueda
            if (filtroreporte_CD)
            {
                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_reporte_concentrado.Rows)
                {
                    row.Visible = true;
                }

                // Restablecer el estado del filtro
                filtroreporte_CD = false;
            }

            string textoBusqueda = txt_filtro_report_consentrado.Text.Trim().ToUpper();
            string semanaSeleccionada = cb_semana_concentrado_d.SelectedItem?.ToString();

            // Validar que al menos un campo de búsqueda esté lleno
            if (string.IsNullOrEmpty(textoBusqueda) && string.IsNullOrEmpty(semanaSeleccionada))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Favor de llenar al menos un campo de búsqueda (texto o semana).",
                    "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Deseleccionar cualquier celda y fila antes de filtrar
            dgv_reporte_concentrado.ClearSelection();
            dgv_reporte_concentrado.CurrentCell = null;

            // Mover el CurrencyManager a una posición válida
            CurrencyManager cm = (CurrencyManager)BindingContext[dgv_reporte_concentrado.DataSource];
            if (cm != null && cm.Count > 0)
                cm.Position = -1;

            bool hayFiltro = false;

            foreach (DataGridViewRow row in dgv_reporte_concentrado.Rows)
            {
                if (row.IsNewRow) continue;

                bool coincideTexto = false;
                bool coincideSemana = false;

                // Aplicar filtro de texto (si hay texto de búsqueda)
                if (!string.IsNullOrEmpty(textoBusqueda))
                {
                    // Buscar en todas las celdas de la fila
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null &&
                            cell.Value.ToString().ToUpper().Contains(textoBusqueda))
                        {
                            coincideTexto = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Si no hay texto de búsqueda, considerar que coincide
                    coincideTexto = true;
                }

                // Aplicar filtro de semana (si hay semana seleccionada)
                if (!string.IsNullOrEmpty(semanaSeleccionada))
                {
                    // Verificar la segunda columna (índice 1) para la semana
                    if (row.Cells[1].Value != null &&
                        row.Cells[1].Value.ToString().Equals(semanaSeleccionada))
                    {
                        coincideSemana = true;
                    }
                }
                else
                {
                    // Si no hay semana seleccionada, considerar que coincide
                    coincideSemana = true;
                }

                // La fila debe coincidir con ambos filtros (si están activos)
                bool mostrarFila = coincideTexto && coincideSemana;

                row.Visible = mostrarFila;
                if (mostrarFila) hayFiltro = true;
            }

            filtroreporte_CD = hayFiltro;

            if (!hayFiltro)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "No se encontraron coincidencias con los criterios de búsqueda.",
                    "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_filtro_consentrado_Click(object sender, EventArgs e)
        {
            try
            {
                // Limpiar el campo de texto de búsqueda
                txt_filtro_report_consentrado.Text = "";

                // Limpiar la selección del ComboBox de semana
                cb_semana_concentrado_d.SelectedIndex = -1;
                cb_semana_concentrado_d.Focus();

                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_reporte_concentrado.Rows)
                {
                    if (!row.IsNewRow) // Asegurarse de no ocultar la fila nueva
                    {
                        row.Visible = true;
                    }
                }

                // Deseleccionar cualquier celda o fila seleccionada
                dgv_reporte_concentrado.ClearSelection();
                dgv_reporte_concentrado.CurrentCell = null;

                // Restablecer el estado del filtro
                filtroreporte_CD = false;

                // Opcional: Restablecer la posición del CurrencyManager si usas binding
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_reporte_concentrado.DataSource];
                if (cm != null && cm.Count > 0)
                {
                    cm.Position = 0; // O -1 dependiendo de tu necesidad
                }

            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al limpiar los filtros: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btn_new_concentrado_otras_Click(object sender, EventArgs e)
        {
            actualiza_reporte_concentrado_otras();
            btn_clean_concentrado_otras.Enabled = true;
            btn_export_excel_concentrado_otras.Enabled = true;
            materialCard13.Enabled = true;
        }
        private void actualiza_reporte_concentrado_otras()
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Obtener el año seleccionado del ComboBox
            string añoSeleccionado = CB_Anio_Con_Otras_A.Text;

            // Validar que se haya seleccionado un año
            if (string.IsNullOrEmpty(añoSeleccionado))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Por favor, seleccione un año antes de generar el reporte de otras áreas.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            //////////////aqui tengo que hace una modificacion las horas efectivas en el reporte solo se restan horas mecanicas de tiempo muerto y esto afecta a todo

            string querySimple = @"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") AS ""Año"",
    TO_CHAR(f.""Fecha"", 'TMMonth') AS ""Mes"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. de Semana"",
    f.""Area"",
    f.""OP"",

    -- AHORA dará 41.08, no 246.48
    SUM(f.""Hr_programadas"") AS ""Hr Programadas"",
    
    -- Horas Reales
    ROUND(SUM(f.""Hr_programadas"") - COALESCE(SUM(tmm.total_min_mecanico) / 60.0, 0), 2) AS ""Hr Reales"",
    
    -- Tiempo Muerto Mecánico (sumado ANTES del JOIN)
    COALESCE(ROUND(SUM(tmm.total_min_mecanico) / 60.0, 2), 0) AS ""Tiempo Muerto Mecánico(Hrs)"",
    
    -- Tiempo Muerto Operativo (sumado ANTES del JOIN)
    COALESCE(ROUND(SUM(tmo.total_min_operativo) / 60.0, 2), 0) AS ""Tiempo Muerto Operativo(Hrs)"",

    -- Porcentaje de cumplimiento
    ROUND(
        (1 - (
            (COALESCE(SUM(tmo.total_min_operativo) / 60.0, 0) + COALESCE(SUM(tmm.total_min_mecanico) / 60.0, 0)) / 
            NULLIF(SUM(f.""Hr_programadas""), 0)
        )) * 100, 
    2) AS ""% Cumplimiento Tiempo Efectivo"",

    -- Kg producidos
    SUM(f.""Kg_prod_term"") AS ""Kg Producidos"",

    -- Kg fuera de especificación
    SUM(f.""Kg_fuera_espec"") AS ""Kg Fuera de Especificación"",

    -- FTT
    ROUND(
        CASE 
            WHEN SUM(f.""Kg_prod_term"") <= 0 THEN 100
            ELSE LEAST(
                ((SUM(f.""Kg_prod_term"") - SUM(f.""Kg_fuera_espec"")) / SUM(f.""Kg_prod_term"")) * 100,
                100
            )
        END,
    2) AS ""FTT"",

    -- Personal promedio
    ROUND(AVG(f.""Personal_Operativo""))::integer AS ""Personal"",

    -- Horas Hombre
    ROUND(
        ROUND(SUM(f.""Hr_programadas"") - COALESCE(SUM(tmm.total_min_mecanico) / 60.0, 0), 2) 
        * ROUND(AVG(f.""Personal_Operativo""))::integer, 
    2) AS ""Horas Hombre"",

    -- Kg producidos por persona
    ROUND(
        CASE 
            WHEN (
                ROUND(SUM(f.""Hr_programadas"") - COALESCE(SUM(tmm.total_min_mecanico) / 60.0, 0), 2) 
                * ROUND(AVG(f.""Personal_Operativo""))::integer
            ) <= 0 THEN 0
            ELSE SUM(f.""Kg_prod_term"") / (
                ROUND(SUM(f.""Hr_programadas"") - COALESCE(SUM(tmm.total_min_mecanico) / 60.0, 0), 2) 
                * ROUND(AVG(f.""Personal_Operativo""))::integer
            )
        END,
    2) AS ""Kg Producidos por Persona"",

    -- Kg de merma
    SUM(f.""Merma_kg"") AS ""Kg de Merma"",

    -- % Merma vs Producido
    ROUND(
        CASE 
            WHEN SUM(f.""Kg_prod_term"") <= 0 THEN 0
            ELSE (SUM(f.""Merma_kg"") / SUM(f.""Kg_prod_term"")) * 100
        END,
    2) AS ""% Merma vs Producido""

FROM public.""Ficha"" f

-- Subconsulta para tiempos muertos operativos (suma ANTES del JOIN)
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_operativo
    FROM public.""Tiempo_Muerto_Operativo""
    GROUP BY ""ID_Ficha""
) tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""

-- Subconsulta para tiempos muertos mecánicos (suma ANTES del JOIN)
LEFT JOIN (
    SELECT 
        ""ID_Ficha"",
        SUM(""Min_Detenidos"") AS total_min_mecanico
    FROM public.""Tiempo_muerto_Mecanico""
    GROUP BY ""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE 
    f.""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
    AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio

GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(MONTH FROM f.""Fecha""),
    TO_CHAR(f.""Fecha"", 'TMMonth'),
    EXTRACT(WEEK FROM f.""Fecha""),
    f.""Area"",
    f.""OP""

ORDER BY 
    ""Año"",
    EXTRACT(MONTH FROM f.""Fecha""),
    ""No. de Semana"",
    f.""Area"",
    f.""OP"";";

            // Crear el parámetro para el año
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
        new NpgsqlParameter("@Anio", NpgsqlTypes.NpgsqlDbType.Integer)
        {
            Value = Convert.ToInt32(añoSeleccionado)
        }
            };

            // Cargar los datos de la tabla Ficha en el DataGridView con el filtro de año
            dbHelper.LoadDataIntoDataGridView(querySimple, dgv_reporte_concentrado_otras, parameters);

            // Configurar el AutoSizeMode de la columna Area (índice 3)
            if (dgv_reporte_concentrado_otras.Columns.Count > 3)
            {
                dgv_reporte_concentrado_otras.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            dgv_reporte_concentrado_otras.ClearSelection();
        }
        //dgv_reporte_concentrado_otras
        private async void btn_export_excel_concentrado_otras_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarDataGridViewFiltradoAExcel_ClosedXML(
                            dgv_reporte_concentrado_otras,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }

        private void btn_clean_concentrado_otras_Click(object sender, EventArgs e)
        {
            dgv_reporte_concentrado_otras.DataSource = null;   // Desvincula cualquier origen de datos
            dgv_reporte_concentrado_otras.Rows.Clear();        // Limpia todas las filas (por si no hay DataSource)
            dgv_reporte_concentrado_otras.Columns.Clear();     // Limpia todas las columnas
            btn_export_excel_concentrado_otras.Enabled = false;
            btn_clean_concentrado_otras.Enabled = false;
            // Limpiar el campo de texto de búsqueda
            txt_filtro_report_consentrado_otras.Text = "";

            // Limpiar la selección del ComboBox de semana
            cb_semana_concentrado_otras.SelectedIndex = -1;
            cb_semana_concentrado_otras.Focus();
            materialCard13.Enabled = false;
        }

        private void btn_buscar_Concen_otras_Click(object sender, EventArgs e)
        {
            // Si el filtro está activo, limpiar filtro y campos de búsqueda
            if (filtroreporte_CO)
            {
                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_reporte_concentrado_otras.Rows)
                {
                    row.Visible = true;
                }

                // Restablecer el estado del filtro
                filtroreporte_CO = false;
            }

            string textoBusqueda = txt_filtro_report_consentrado_otras.Text.Trim().ToUpper();
            string semanaSeleccionada = cb_semana_concentrado_otras.SelectedItem?.ToString();

            // Validar que al menos un campo de búsqueda esté lleno
            if (string.IsNullOrEmpty(textoBusqueda) && string.IsNullOrEmpty(semanaSeleccionada))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Favor de llenar al menos un campo de búsqueda (texto o semana).",
                    "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Deseleccionar cualquier celda y fila antes de filtrar
            dgv_reporte_concentrado_otras.ClearSelection();
            dgv_reporte_concentrado_otras.CurrentCell = null;

            // Mover el CurrencyManager a una posición válida
            CurrencyManager cm = (CurrencyManager)BindingContext[dgv_reporte_concentrado_otras.DataSource];
            if (cm != null && cm.Count > 0)
                cm.Position = -1;

            bool hayFiltro = false;

            foreach (DataGridViewRow row in dgv_reporte_concentrado_otras.Rows)
            {
                if (row.IsNewRow) continue;

                bool coincideTexto = false;
                bool coincideSemana = false;

                // Aplicar filtro de texto (si hay texto de búsqueda)
                if (!string.IsNullOrEmpty(textoBusqueda))
                {
                    // Buscar en todas las celdas de la fila
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null &&
                            cell.Value.ToString().ToUpper().Contains(textoBusqueda))
                        {
                            coincideTexto = true;
                            break;
                        }
                    }
                }
                else
                {
                    // Si no hay texto de búsqueda, considerar que coincide
                    coincideTexto = true;
                }

                // Aplicar filtro de semana (si hay semana seleccionada)
                if (!string.IsNullOrEmpty(semanaSeleccionada))
                {
                    // Verificar la segunda columna (índice 1) para la semana
                    if (row.Cells[2].Value != null &&
                        row.Cells[2].Value.ToString().Equals(semanaSeleccionada))
                    {
                        coincideSemana = true;
                    }
                }
                else
                {
                    // Si no hay semana seleccionada, considerar que coincide
                    coincideSemana = true;
                }

                // La fila debe coincidir con ambos filtros (si están activos)
                bool mostrarFila = coincideTexto && coincideSemana;

                row.Visible = mostrarFila;
                if (mostrarFila) hayFiltro = true;
            }

            filtroreporte_CO = hayFiltro;

            if (!hayFiltro)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "No se encontraron coincidencias con los criterios de búsqueda.",
                    "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_filtro_consentrado_otras_Click(object sender, EventArgs e)
        {
            try
            {
                // Limpiar el campo de texto de búsqueda
                txt_filtro_report_consentrado_otras.Text = "";

                // Limpiar la selección del ComboBox de semana
                cb_semana_concentrado_otras.SelectedIndex = -1;
                cb_semana_concentrado_otras.Focus();

                // Mostrar todas las filas del DataGridView
                foreach (DataGridViewRow row in dgv_reporte_concentrado_otras.Rows)
                {
                    if (!row.IsNewRow) // Asegurarse de no ocultar la fila nueva
                    {
                        row.Visible = true;
                    }
                }

                // Deseleccionar cualquier celda o fila seleccionada
                dgv_reporte_concentrado_otras.ClearSelection();
                dgv_reporte_concentrado_otras.CurrentCell = null;

                // Restablecer el estado del filtro
                filtroreporte_CO = false;

                // Opcional: Restablecer la posición del CurrencyManager si usas binding
                CurrencyManager cm = (CurrencyManager)BindingContext[dgv_reporte_concentrado_otras.DataSource];
                if (cm != null && cm.Count > 0)
                {
                    cm.Position = 0; // O -1 dependiendo de tu necesidad
                }

            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al limpiar los filtros: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void bnt_graficar_Click(object sender, EventArgs e)
        {
            // Validar que se haya seleccionado un tipo de gráfica
            if (cb_tipo_grafica.SelectedIndex == -1)
            {
                lista_semanas.Visible = false; // Ocultar la lista de semanas
                materialCard26.Visible = false; // Ocultar el panel de gráficos
                MetroFramework.MetroMessageBox.Show(this, "Por favor seleccione un tipo de gráfica.", "Tipo requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validar que se haya seleccionado un año
            if (CB_Anio_grafica.SelectedItem == null)
            {
                lista_semanas.Visible = false; // Ocultar la lista de semanas
                materialCard26.Visible = false; // Ocultar el panel de gráficos
                MetroFramework.MetroMessageBox.Show(this, "Por favor seleccione un año para graficar.", "Año requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ÍNDICES DE GRÁFICAS QUE NO REQUIEREN SEMANAS
            int[] graficasSinSemanas = { 8, 10 }; // Case 8 = Cumplimiento Mensual

            // Si la gráfica NO requiere semanas, llamar directamente
            if (graficasSinSemanas.Contains(cb_tipo_grafica.SelectedIndex))
            {
                switch (cb_tipo_grafica.SelectedIndex)
                {
                    case 8: // Cumplimiento Mensual
                        GraficarCumplimientoMensualPorSupervisor();
                        materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                        break;
                    case 10: // Cumplimiento Mensual otras áreas
                        GraficarCumplimientoMensualPorSupervisorOtrasAreas();
                        materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                        break;
                }
                return; // Importante: salir después de graficar
            }

            // PARA LAS DEMÁS GRÁFICAS: Validar semanas seleccionadas
            var seleccionados = lista_semanas.CheckedItems;
            if (seleccionados.Count == 0)
            {
                materialCard26.Visible = false;
                MetroFramework.MetroMessageBox.Show(this,
                    "Favor de seleccionar una o varias semanas de la lista para continuar.",
                    "No hay semanas seleccionadas.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // Crear lista de semanas seleccionadas
            var semanasSeleccionadas = new List<string>();
            foreach (ListViewDataItem item in seleccionados)
            {
                semanasSeleccionadas.Add(item.Text);
            }

            // Llamar a la gráfica correspondiente según el combo box
            switch (cb_tipo_grafica.SelectedIndex)
            {
                case 0: // Kg Fresco
                    GraficarKgFresco(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 1: // Kg Seco
                    GraficarKgSeco(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 2: // FTT Deshidratado
                    GraficarFTTDeshidratado(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 3: // FTT Otras Areas
                    GraficarFTTOtrasAreas(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 4: // Merma por Supervisor
                    GraficarMermasPorSupervisor(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 5: // Cumplimiento Despegue
                    GraficarCumplimientoDespegue(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 6: // Cumplimiento Otras Areas
                    GraficarCumplimientoOtrasAreas(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                case 7: // Cumplimiento por Supervisor (Semanal)
                    GraficarCumplimientoPorSupervisor(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;
                case 9: // Cumplimiento por Supervisor (Semanal)
                    GraficarCumplimientoPorSupervisorOtrasAreas(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;
                case 11: // Cumplimiento por Supervisor (Semanal)
                    GraficarCumplimientoPorJefe(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;
                case 12: // Cumplimiento por Supervisor (Semanal)
                    GraficarCumplimientoPorJefeOtrasAreas(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;
                case 13: // Cumplimiento por Supervisor (Semanal)
                    GraficarCumplimientoGeneral(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;
                case 14: // Tiempo Efectivo (Semanal)
                    GraficarTiempoEfectivo(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;
                case 15: // Tiempo Efectivo (Semanal)
                    GraficarTiempoEfectivoProduccion(semanasSeleccionadas);
                    materialCard26.Visible = true; // Asegurar que el panel de gráficos esté visible
                    break;

                default:
                    materialCard26.Visible = false; // Ocultar el panel de gráficos en caso de tipo no válido
                    MetroFramework.MetroMessageBox.Show(this, "Tipo de gráfica no válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarTiempoEfectivoProduccion(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% TIEMPO EFECTIVO PRODUCCION - ANÁLISIS SEMANAL",
                    "Porcentaje de Tiempo Efectivo - Todas las Áreas excluyendo Deshidratado"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;
                string query = @"
        SELECT 
            EXTRACT(YEAR FROM f.""Fecha"") AS ""Año"",
            EXTRACT(WEEK FROM f.""Fecha"") AS ""No_Semana"",
            CEIL(
                AVG(
                    (1 - (
                        (COALESCE(tmo_daily.total_min_detenidos/60, 0) + COALESCE(tmm_daily.total_min_detenidos/60, 0)) / 
                        NULLIF(f.""Hr_programadas"", 0)
                    ))*100
                )) AS ""Promedio_Cumplimiento_Tiempo_Efectivo""
        FROM 
            public.""Ficha"" f
        LEFT JOIN (
            SELECT 
                ""ID_Ficha"",
                SUM(""Min_Detenidos"")/60 AS total_min_detenidos
            FROM 
                public.""Tiempo_Muerto_Operativo""
            GROUP BY 
                ""ID_Ficha""
        ) tmo_daily ON f.""ID_Ficha"" = tmo_daily.""ID_Ficha""
        LEFT JOIN (
            SELECT 
                ""ID_Ficha"",
                SUM(""Min_Detenidos"")/60 AS total_min_detenidos
            FROM 
                public.""Tiempo_muerto_Mecanico""
            GROUP BY 
                ""ID_Ficha""
        ) tmm_daily ON f.""ID_Ficha"" = tmm_daily.""ID_Ficha""
        WHERE 
            f.""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
            AND f.""Hr_programadas"" > 0
            AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
            AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
        GROUP BY 
            EXTRACT(YEAR FROM f.""Fecha""),
            EXTRACT(WEEK FROM f.""Fecha"")
        ORDER BY 
            ""Año"", ""No_Semana""";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarTiempoEfectivoProduccion(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarTiempoEfectivoProduccion(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresCumplimiento = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                double cumplimiento = System.Convert.ToDouble(row["Promedio_Cumplimiento_Tiempo_Efectivo"]);
                valoresCumplimiento.Add(cumplimiento);
                semanas.Add($"Sem {row["No_Semana"]}");
            }

            // Crear serie de línea para FTT - CORREGIDO
            var serieFTT = new LiveCharts.Wpf.LineSeries
            {
                Title = "% Tiempo Efectivo",
                Values = valoresCumplimiento,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 152, 219)), // Azul claro
                StrokeThickness = 4,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 14,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(41, 128, 185)), // Color del punto (relleno)
                                                                                                                              // ❌ NO USAR PointBackground (no existe en LiveCharts 1.x)
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                LineSmoothness = 0, // Ligeramente suavizada
            };

            // Asignar series (solo una para FTT)
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieFTT
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "PORCENTAJE (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // FTT generalmente va de 0 a 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarTiempoEfectivo(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% TIEMPO EFECTIVO - ANÁLISIS SEMANAL",
                    "Porcentaje de Tiempo Efectivo - Área Deshidratado"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;
                string query = @"
        SELECT 
            EXTRACT(YEAR FROM f.""Fecha"") AS ""Año"",
            EXTRACT(WEEK FROM f.""Fecha"") AS ""No_Semana"",
            ROUND(AVG(
                CASE 
                    WHEN f.""Hr_programadas"" > 0 THEN
                        (1 -((COALESCE(tmo_avg.""Avg_Operativo"", 0) + COALESCE(tmm_avg.""Avg_Mecanico"", 0)) / f.""Hr_programadas""))*100 
                    ELSE 0
                END
            ), 2) AS ""Promedio_Cumplimiento_Tiempo_Efectivo""
        FROM public.""Ficha"" f
        LEFT JOIN (
            SELECT 
                tmo.""ID_Ficha"",
                SUM(tmo.""Min_Detenidos""/60) AS ""Avg_Operativo""
            FROM public.""Tiempo_Muerto_Operativo"" tmo
            GROUP BY tmo.""ID_Ficha""
        ) tmo_avg ON f.""ID_Ficha"" = tmo_avg.""ID_Ficha""
        LEFT JOIN (
            SELECT 
                tmm.""ID_Ficha"",
                SUM(tmm.""Min_Detenidos""/60) AS ""Avg_Mecanico""
            FROM public.""Tiempo_muerto_Mecanico"" tmm
            GROUP BY tmm.""ID_Ficha""
        ) tmm_avg ON f.""ID_Ficha"" = tmm_avg.""ID_Ficha""
        WHERE f.""Area"" IN ('Despegue')
            AND f.""Hr_programadas"" > 0
            AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
            AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
        GROUP BY EXTRACT(YEAR FROM f.""Fecha""), EXTRACT(WEEK FROM f.""Fecha"")
        ORDER BY ""Año"", ""No_Semana""";


                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarTiempoEfectivo(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarTiempoEfectivo(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresCumplimiento = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                double cumplimiento = System.Convert.ToDouble(row["Promedio_Cumplimiento_Tiempo_Efectivo"]);
                valoresCumplimiento.Add(cumplimiento);
                semanas.Add($"Sem {row["No_Semana"]}");
            }

            // Crear serie de línea para FTT - CORREGIDO
            var serieFTT = new LiveCharts.Wpf.LineSeries
            {
                Title = "% Tiempo Efectivo",
                Values = valoresCumplimiento,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 152, 219)), // Azul claro
                StrokeThickness = 4,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 14,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(41, 128, 185)), // Color del punto (relleno)
                                                                                                                              // ❌ NO USAR PointBackground (no existe en LiveCharts 1.x)
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                LineSmoothness = 0, // Ligeramente suavizada
            };

            // Asignar series (solo una para FTT)
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieFTT
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "PORCENTAJE (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // FTT generalmente va de 0 a 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarCumplimientoGeneral(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO GENERAL - ANÁLISIS SEMANAL",
                    "Porcentaje de cumplimiento"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;
                string query = $@"
WITH 
despegue AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") as No_Semana,
        ROUND(LEAST(
            (SUM(""Kg_prod_seco"") - SUM(""Kg_fuera_espec"")) / NULLIF(SUM(""Kg_meta""), 0) * 100,
            100
        ), 2) as despegue_valor
    FROM public.""Ficha""
    WHERE ""Area"" = 'Despegue'
        AND EXTRACT(YEAR FROM ""Fecha"") = {añoSeleccionado}
        AND EXTRACT(WEEK FROM ""Fecha"") IN ({semanasParam})
    GROUP BY EXTRACT(WEEK FROM ""Fecha"")
),
otras_areas AS (
    SELECT 
        EXTRACT(WEEK FROM ""Fecha"") as No_Semana,
        ROUND(LEAST(
            (SUM(""Kg_prod_term"") - SUM(""Kg_fuera_espec"")) / NULLIF(SUM(""Kg_meta""), 0) * 100,
            100
        ), 2) as otras_valor
    FROM public.""Ficha""
    WHERE ""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
        AND EXTRACT(YEAR FROM ""Fecha"") = {añoSeleccionado}
        AND EXTRACT(WEEK FROM ""Fecha"") IN ({semanasParam})
    GROUP BY EXTRACT(WEEK FROM ""Fecha"")
)
SELECT 
    COALESCE(d.No_Semana, o.No_Semana) as No_Semana,
    ROUND(
        (COALESCE(d.despegue_valor, 0) + COALESCE(o.otras_valor, 0)) / 2,
        2
    ) as ""% Cumplimiento General""
FROM despegue d
FULL OUTER JOIN otras_areas o ON d.No_Semana = o.No_Semana
ORDER BY No_Semana;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoGeneral(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoGeneral(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresCumplimiento = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento General"]);
                valoresCumplimiento.Add(cumplimiento);
                semanas.Add($"Sem {row["No_Semana"]}");
            }

            // Crear serie de línea para FTT - CORREGIDO
            var serieFTT = new LiveCharts.Wpf.LineSeries
            {
                Title = "% Cumplimiento de Producción",
                Values = valoresCumplimiento,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 152, 219)), // Azul claro
                StrokeThickness = 4,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 14,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(41, 128, 185)), // Color del punto (relleno)
                                                                                                                              // ❌ NO USAR PointBackground (no existe en LiveCharts 1.x)
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                LineSmoothness = 0, // Ligeramente suavizada
            };

            // Asignar series (solo una para FTT)
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieFTT
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "PORCENTAJE (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // FTT generalmente va de 0 a 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarCumplimientoPorJefeOtrasAreas(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO POR JEFE DE TURNO - ANÁLISIS SEMANAL",
                    "Porcentaje de cumplimiento(Excluyendo Despegue)"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;
                string query = @"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") as año,
    EXTRACT(WEEK FROM f.""Fecha"") as No_Semana,
    u.""Usuario"" as ""Jefe de Turno"",  -- Sin COALESCE, solo jefes reales

    -- Porcentaje limitado al 100%
    ROUND(LEAST(
        (SUM(f.""Kg_prod_term"") - SUM(f.""Kg_fuera_espec"")) / NULLIF(SUM(f.""Kg_meta""), 0) * 100,
        100
    ), 2) as ""% Cumplimiento""

FROM public.""Ficha"" f
INNER JOIN public.""Usuarios"" u ON f.""ID_Jefe"" = u.""ID_User""  -- Cambiado a INNER JOIN
WHERE f.""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
    AND u.""Nivel"" = 'Jefe de Turno'  -- Filtro por nivel Jefe de Turno
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(WEEK FROM f.""Fecha""),
    u.""Usuario""
ORDER BY 
    año,
    No_Semana,
    ""Jefe de Turno"";";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoPorJefeOtrasAreas(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoPorJefeOtrasAreas(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y semanas únicas
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Jefe de Turno"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            var semanas = datos.AsEnumerable()
                               .Select(row => $"Sem {row.Field<decimal>("No_Semana")}")
                               .Distinct()
                               .OrderBy(s => s)
                               .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todas las semanas con 0
                foreach (var semana in semanas)
                {
                    valoresPorSupervisor[supervisor][semana] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Jefe de Turno"].ToString();
                string semana = $"Sem {row["No_Semana"]}";
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][semana] = cumplimiento;
                }
            }

            // Array de colores para los supervisores (colores vibrantes para líneas)
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(52, 152, 219),  // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),  // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),   // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),  // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),  // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),  // Naranja
        System.Windows.Media.Color.FromRgb(26, 188, 156),  // Turquesa
        System.Windows.Media.Color.FromRgb(52, 73, 94),    // Azul oscuro
        System.Windows.Media.Color.FromRgb(192, 57, 43),   // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),  // Azul medio
            };

            // Crear una serie de línea por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var semana in semanas)
                {
                    valores.Add(valoresPorSupervisor[supervisor][semana]);
                }

                var color = colores[colorIndex % colores.Length];

                var serie = new LiveCharts.Wpf.LineSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Stroke = new System.Windows.Media.SolidColorBrush(color),
                    StrokeThickness = 3,
                    Fill = System.Windows.Media.Brushes.Transparent,
                    PointGeometrySize = 12,
                    PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                    PointForeground = System.Windows.Media.Brushes.White,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                    LineSmoothness = 0, // Suavizado moderado
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "CUMPLIMIENTO (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // Cumplimiento limitado al 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // Agregar una línea horizontal en el 100% (opcional - meta ideal)
            // Nota: LiveCharts 1.x no tiene líneas de anotación fáciles, 
            // pero podemos agregar un separador especial en el eje Y

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarCumplimientoPorJefe(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO POR JEFE DE TURNO - ANÁLISIS SEMANAL",
                    "Porcentaje de cumplimiento de Despegue"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;
                string query = @"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") as año,
    EXTRACT(WEEK FROM f.""Fecha"") as No_Semana,
    u.""Usuario"" as ""Jefe de Turno"",  -- Sin COALESCE, solo jefes reales

    -- Porcentaje limitado al 100%
    ROUND(LEAST(
        (SUM(f.""Kg_prod_seco"") - SUM(f.""Kg_fuera_espec"")) / NULLIF(SUM(f.""Kg_meta""), 0) * 100,
        100
    ), 2) as ""% Cumplimiento""

FROM public.""Ficha"" f
INNER JOIN public.""Usuarios"" u ON f.""ID_Jefe"" = u.""ID_User""  -- Cambiado a INNER JOIN
WHERE f.""Area"" = 'Despegue'
    AND u.""Nivel"" = 'Jefe de Turno'  -- Filtro por nivel Jefe de Turno
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(WEEK FROM f.""Fecha""),
    u.""Usuario""
ORDER BY 
    año,
    No_Semana,
    ""Jefe de Turno"";";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoPorJefe(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoPorJefe(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y semanas únicas
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Jefe de Turno"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            var semanas = datos.AsEnumerable()
                               .Select(row => $"Sem {row.Field<decimal>("No_Semana")}")
                               .Distinct()
                               .OrderBy(s => s)
                               .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todas las semanas con 0
                foreach (var semana in semanas)
                {
                    valoresPorSupervisor[supervisor][semana] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Jefe de Turno"].ToString();
                string semana = $"Sem {row["No_Semana"]}";
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][semana] = cumplimiento;
                }
            }

            // Array de colores para los supervisores (colores vibrantes para líneas)
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(52, 152, 219),  // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),  // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),   // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),  // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),  // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),  // Naranja
        System.Windows.Media.Color.FromRgb(26, 188, 156),  // Turquesa
        System.Windows.Media.Color.FromRgb(52, 73, 94),    // Azul oscuro
        System.Windows.Media.Color.FromRgb(192, 57, 43),   // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),  // Azul medio
            };

            // Crear una serie de línea por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var semana in semanas)
                {
                    valores.Add(valoresPorSupervisor[supervisor][semana]);
                }

                var color = colores[colorIndex % colores.Length];

                var serie = new LiveCharts.Wpf.LineSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Stroke = new System.Windows.Media.SolidColorBrush(color),
                    StrokeThickness = 3,
                    Fill = System.Windows.Media.Brushes.Transparent,
                    PointGeometrySize = 12,
                    PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                    PointForeground = System.Windows.Media.Brushes.White,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                    LineSmoothness = 0, // Suavizado moderado
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "CUMPLIMIENTO (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // Cumplimiento limitado al 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // Agregar una línea horizontal en el 100% (opcional - meta ideal)
            // Nota: LiveCharts 1.x no tiene líneas de anotación fáciles, 
            // pero podemos agregar un separador especial en el eje Y

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        private void GraficarKgFresco(List<string> semanasSeleccionadas)
        {
            try
            {
                ConfigurarTitulosGrafico("Kilogramos en Fresco que Ingresan a Túnel", "Comparación Meta Programada / Meta de Horas Reales Trabajadas / Real Ingresado a Túnel");
                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;
                ////////////////////////////horas programas - tmm
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////sume el tiempo operativo en Hr_efectivas
                string query = @"
SELECT 
    EXTRACT(WEEK FROM f.""Fecha"") AS semana,
    SUM(f.""Hr_programadas"" * d.""Kg_fresco_hr"") AS ""Suma_Kg_Fresco_Meta_por_Hr_programadas"",
    
    -- Horas Reales = Horas Efectivas - Tiempo Muerto Mecánico (convertido a horas)
    SUM((f.""Hr_programadas"" - COALESCE(tmm.""Total_Horas_Muerto_Mecanico"", 0)) * d.""Kg_fresco_hr"") AS ""Kg_Fresco_Meta_por_Hr_Reales"",
    
    SUM(f.kg_frescos_enter_se) AS ""Kg_Fresco_Real""
FROM 
    public.""Ficha"" f
INNER JOIN 
    public.""Deshidratado"" d ON f.""OP"" = d.""OP""
LEFT JOIN (
    -- Subconsulta para calcular el total de minutos muertos mecánicos por ficha
    SELECT 
        tmm.""ID_Ficha"",
        SUM(tmm.""Min_Detenidos"") / 60.0 AS ""Total_Horas_Muerto_Mecanico""
    FROM 
        public.""Tiempo_muerto_Mecanico"" tmm
    GROUP BY 
        tmm.""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE 
    f.""Area"" = 'Tunel/Sumergidor'
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
GROUP BY 
    EXTRACT(WEEK FROM f.""Fecha"")
ORDER BY 
    semana";
                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "No se encontraron datos para las semanas seleccionadas.", "Precaución", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoKgFresco(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, $"Error al graficar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConfigurarGraficoKgFresco(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresMetaProgramada = new LiveCharts.ChartValues<double>();
            var valoresMetaReales = new LiveCharts.ChartValues<double>();
            var valoresReal = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                valoresMetaProgramada.Add(System.Convert.ToDouble(row["Suma_Kg_Fresco_Meta_por_Hr_programadas"]));
                valoresMetaReales.Add(System.Convert.ToDouble(row["Kg_Fresco_Meta_por_Hr_Reales"]));
                valoresReal.Add(System.Convert.ToDouble(row["Kg_Fresco_Real"]));
                semanas.Add($"Sem {row["semana"]}");
            }

            // Crear series con títulos para la leyenda
            var serieMetaProgramada = new LiveCharts.Wpf.ColumnSeries
            {
                Title = "Meta Programada", // ← Esto es lo que aparece en la leyenda
                Values = valoresMetaProgramada,
                Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 134, 232)),
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(45, 95, 175)),
                StrokeThickness = 1,
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("N0"),
                MaxColumnWidth = 50,
            };

            var serieMetaReales = new LiveCharts.Wpf.ColumnSeries
            {
                Title = "Meta Reales", // ← Esto es lo que aparece en la leyenda
                Values = valoresMetaReales,
                Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(46, 204, 113)),
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(35, 155, 86)),
                StrokeThickness = 1,
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("N0"),
                MaxColumnWidth = 50,
            };

            var serieReal = new LiveCharts.Wpf.LineSeries
            {
                Title = "Real", // ← Esto es lo que aparece en la leyenda
                Values = valoresReal,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(231, 76, 60)),
                StrokeThickness = 3,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 12,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("N0"),
                LineSmoothness = 0,
            };

            // Asignar series
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieMetaProgramada,
        serieMetaReales,
        serieReal
    };

            // CONFIGURACIÓN DE LA LEYENDA EN LIVECHARTS 1.x
            // Solamente necesitas esto para que aparezca la leyenda:
            cartesianChart1.LegendLocation = LegendLocation.Bottom; // Bottom, Top, Left, Right
            cartesianChart1.Hoverable = true; // Efectos al pasar el mouse

            // ⚠️ NOTA: ChartLegend NO existe en LiveCharts 1.x
            // La personalización de la leyenda es limitada en esta versión

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator
                {
                    IsEnabled = false
                }
            });

            // Eje Y
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "KILOGRAMOS",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("N0"),
                MinValue = 0,
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null; // Quita el toolbar flotante
        }
        /// <summary>
        /// Grafica de Kilogramos en Seco en Deshidratador
        /// </summary>
        private void GraficarKgSeco(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Kg Seco
                ConfigurarTitulosGrafico(
                    "Kilogramos en Secos en Deshidratador",
                    "Comparativo entre Meta según hras programadas / Meta según hras reales trabajadas / Producido"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = @"
SELECT 
    EXTRACT(WEEK FROM f.""Fecha"") AS semana,
    SUM(f.""Hr_programadas"" * d.""Kg_seco_hr"") AS ""Suma_Kg_Seco_Meta_por_Hr_programadas"",
    
    -- Kg Seco Meta por Horas Reales = (Horas Efectivas - Tiempo Muerto Mecánico) * Kg_seco_hr
    SUM((f.""Hr_programadas"" - COALESCE(tmm.""Total_Horas_Muerto_Mecanico"", 0)) * d.""Kg_seco_hr"") AS ""Kg_Seco_Meta_por_Hr_Reales"",
    
    SUM(f.""Kg_prod_seco"") AS ""Kg_Seco_Real""
FROM 
    public.""Ficha"" f
INNER JOIN 
    public.""Deshidratado"" d ON f.""OP"" = d.""OP""
LEFT JOIN (
    -- Subconsulta para calcular el total de minutos muertos mecánicos por ficha
    SELECT 
        tmm.""ID_Ficha"",
        SUM(tmm.""Min_Detenidos"") / 60.0 AS ""Total_Horas_Muerto_Mecanico""
    FROM 
        public.""Tiempo_muerto_Mecanico"" tmm
    GROUP BY 
        tmm.""ID_Ficha""
) tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""

WHERE 
    f.""Area"" = 'Despegue'
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
GROUP BY 
    EXTRACT(WEEK FROM f.""Fecha"")
ORDER BY 
    semana";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en Kg Seco.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts (REUTILIZANDO el mismo método pero con los datos de Kg Seco)
                ConfigurarGraficoKgSeco(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar Kg Seco: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para Kg Seco
        /// </summary>
        private void ConfigurarGraficoKgSeco(DataTable datos)
        {
            // Limpiar el chart de LiveCharts (REUTILIZANDO el mismo cartesianChart1)
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresMetaProgramada = new LiveCharts.ChartValues<double>();
            var valoresMetaReales = new LiveCharts.ChartValues<double>();
            var valoresReal = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                valoresMetaProgramada.Add(System.Convert.ToDouble(row["Suma_Kg_Seco_Meta_por_Hr_programadas"]));
                valoresMetaReales.Add(System.Convert.ToDouble(row["Kg_Seco_Meta_por_Hr_Reales"]));
                valoresReal.Add(System.Convert.ToDouble(row["Kg_Seco_Real"]));
                semanas.Add($"Sem {row["semana"]}");
            }

            // Crear series con títulos para la leyenda
            var serieMetaProgramada = new LiveCharts.Wpf.ColumnSeries
            {
                Title = "Meta Programada",
                Values = valoresMetaProgramada,
                Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(74, 134, 232)), // Azul
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(45, 95, 175)),
                StrokeThickness = 1,
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("N0"),
                MaxColumnWidth = 50,
            };

            var serieMetaReales = new LiveCharts.Wpf.ColumnSeries
            {
                Title = "Meta Reales",
                Values = valoresMetaReales,
                Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(46, 204, 113)), // Verde
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(35, 155, 86)),
                StrokeThickness = 1,
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("N0"),
                MaxColumnWidth = 50,
            };

            var serieReal = new LiveCharts.Wpf.LineSeries
            {
                Title = "Real",
                Values = valoresReal,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(231, 76, 60)), // Rojo
                StrokeThickness = 3,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 12,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("N0"),
                LineSmoothness = 0,
            };

            // Asignar series
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieMetaProgramada,
        serieMetaReales,
        serieReal
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "KILOGRAMOS",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("N0"),
                MinValue = 0,
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de FTT Deshidratado - Análisis Semanal
        /// </summary>
        private void GraficarFTTDeshidratado(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para FTT
                ConfigurarTitulosGrafico(
                    "% FTT DESHIDRATADO - ANÁLISIS SEMANAL",
                    "First Time Through"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = @"
SELECT 
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No_Semana"",
    
    -- Promedio de FTT multiplicado por 100 (redondeado hacia arriba)
    CEIL(AVG(f.""FTT"") * 100) AS ""FTT""
    
FROM 
    public.""Ficha"" f
WHERE 
    f.""Area"" = 'Despegue'
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
GROUP BY 
    EXTRACT(WEEK FROM f.""Fecha"")
ORDER BY 
    ""No_Semana"";";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en FTT Deshidratado.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoFTT(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar FTT Deshidratado: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Grafica de FTT Deshidratado - Análisis Semanal
        /// </summary>
        private void GraficarFTTOtrasAreas(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para FTT
                ConfigurarTitulosGrafico(
                    "% FTT PROMEDIO OTRAS ÁREAS - ANÁLISIS SEMANAL",
                    "First Time Through - Promedio de todas las áreas de producción"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = @"SELECT 
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No_Semana"",
    
    -- Fórmula FTT: (Kg_prod_term - Kg_fuera_espec) / Kg_prod_term * 100
    CASE 
        WHEN SUM(f.""Kg_prod_term"") <= 0 THEN 100
        ELSE CEIL(
		LEAST(
            ((SUM(f.""Kg_prod_term"") - SUM(f.""Kg_fuera_espec"")) / SUM(f.""Kg_prod_term"")) * 100,
            100
        )
		)
    END AS ""FTT""
    
FROM 
    public.""Ficha"" f
WHERE 
    f.""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
GROUP BY 
    EXTRACT(WEEK FROM f.""Fecha"")
ORDER BY 
    ""No_Semana"";";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en FTT",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoFTT(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar FTT: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para FTT Deshidratado (Línea)
        /// </summary>
        private void ConfigurarGraficoFTT(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresFTT = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                double ftt = System.Convert.ToDouble(row["FTT"]);
                valoresFTT.Add(ftt);
                semanas.Add($"Sem {row["No_Semana"]}");
            }

            // Crear serie de línea para FTT - CORREGIDO
            var serieFTT = new LiveCharts.Wpf.LineSeries
            {
                Title = "% FTT",
                Values = valoresFTT,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 152, 219)), // Azul claro
                StrokeThickness = 4,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 14,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(41, 128, 185)), // Color del punto (relleno)
                                                                                                                              // ❌ NO USAR PointBackground (no existe en LiveCharts 1.x)
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                LineSmoothness = 0, // Ligeramente suavizada
            };

            // Asignar series (solo una para FTT)
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieFTT
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "PORCENTAJE (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // FTT generalmente va de 0 a 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
           // cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de Mermas por Supervisor
        /// </summary>
        private void GraficarMermasPorSupervisor(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Mermas
                ConfigurarTitulosGrafico(
                    "KILOGRAMOS DE MERMAS POR SUPERVISOR - ANÁLISIS SEMANAL",
                    "Comparativo de merma generada por cada supervisor"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = @"
WITH merma_usuario AS (
    SELECT 
        u.""Usuario"" as ""Nombre_Usuario"",
        EXTRACT(WEEK FROM f.""Fecha"") AS numero_semana, 
        EXTRACT(YEAR FROM f.""Fecha"") AS año,
        SUM(
            CASE 
                WHEN f.""Area"" = 'Tunel/Sumergidor' THEN
                    f.""Merma_podrido"" + f.""Merma_tina"" + f.""Merma_piso"" + f.""Merma_canaletas"" + f.""Merma_lavado_bandas""
                ELSE f.""Merma_kg""
            END
        ) AS total_merma_kg
    FROM public.""Ficha"" f
    INNER JOIN public.""Usuarios"" u ON f.""ID_user"" = u.""ID_User""
    WHERE f.""Area"" IS NOT NULL
        AND u.""Nivel"" = 'Supervisor'  -- Filtro por Nivel Supervisor
        AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
        AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
    GROUP BY u.""Usuario"", EXTRACT(WEEK FROM f.""Fecha""), EXTRACT(YEAR FROM f.""Fecha"")
)

SELECT 
    ""Nombre_Usuario"" AS ""Supervisor"",
    numero_semana as ""No_Semana"",
    año as ""Año"",
    COALESCE(total_merma_kg, 0) as ""Merma_Kg""
FROM merma_usuario
ORDER BY año, numero_semana, ""Nombre_Usuario"";";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en Mermas por Supervisor.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoMermasPorSupervisor(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar Mermas por Supervisor: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para Mermas por Supervisor (Barras agrupadas)
        /// </summary>
        private void ConfigurarGraficoMermasPorSupervisor(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y semanas únicas
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Supervisor"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            var semanas = datos.AsEnumerable()
                               .Select(row => $"Sem {row.Field<decimal>("No_Semana")}")
                               .Distinct()
                               .OrderBy(s => s)
                               .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todas las semanas con 0
                foreach (var semana in semanas)
                {
                    valoresPorSupervisor[supervisor][semana] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Supervisor"].ToString();
                string semana = $"Sem {row["No_Semana"]}";
                double merma = System.Convert.ToDouble(row["Merma_Kg"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][semana] = merma;
                }
            }

            // Array de colores para los supervisores
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(74, 134, 232),   // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),   // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),    // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),   // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),   // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),   // Naranja
        System.Windows.Media.Color.FromRgb(52, 73, 94),     // Azul oscuro
        System.Windows.Media.Color.FromRgb(22, 160, 133),   // Verde azulado
        System.Windows.Media.Color.FromRgb(192, 57, 43),    // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),   // Azul medio
            };

            // Crear una serie por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var semana in semanas)
                {
                    valores.Add(valoresPorSupervisor[supervisor][semana]);
                }

                var color = colores[colorIndex % colores.Length];
                var colorOscuro = System.Windows.Media.Color.FromRgb(
                    (byte)(color.R * 0.7),
                    (byte)(color.G * 0.7),
                    (byte)(color.B * 0.7)
                );

                var serie = new LiveCharts.Wpf.ColumnSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Fill = new System.Windows.Media.SolidColorBrush(color),
                    Stroke = new System.Windows.Media.SolidColorBrush(colorOscuro),
                    StrokeThickness = 1,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("N0"), // Formato sin decimales
                    MaxColumnWidth = 40,
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "MERMAS (Kg)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("N0"),
                MinValue = 0,
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
           // cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de FTT Deshidratado - Análisis Semanal
        /// </summary>
        private void GraficarCumplimientoDespegue(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para FTT
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO DE PRODUCCIÓN - DESPEGUE",
                    "Cumplimiento semanal de producción en el área de Despegue"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = $@"
        SELECT 
            EXTRACT(YEAR FROM ""Fecha"") as año,
            EXTRACT(WEEK FROM ""Fecha"") as No_Semana,

            -- Porcentaje limitado al 100%
            ROUND(LEAST(
                (SUM(""Kg_prod_seco"") - SUM(""Kg_fuera_espec"")) / NULLIF(SUM(""Kg_meta""), 0) * 100,
                100
            ),2) as ""% Cumplimiento""
        FROM public.""Ficha""
        WHERE ""Area"" = 'Despegue'
            AND EXTRACT(YEAR FROM ""Fecha"") = {añoSeleccionado}
            AND EXTRACT(WEEK FROM ""Fecha"") IN ({semanasParam})
        GROUP BY 
            EXTRACT(YEAR FROM ""Fecha""),
            EXTRACT(WEEK FROM ""Fecha"")
        ORDER BY 
            año,
            No_Semana;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en Cumplimiento",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarCumplimientoDespegue(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para FTT Deshidratado (Línea)
        /// </summary>
        private void ConfigurarCumplimientoDespegue(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresCumplimiento = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);
                valoresCumplimiento.Add(cumplimiento);
                semanas.Add($"Sem {row["No_Semana"]}");
            }

            // Crear serie de línea para FTT - CORREGIDO
            var serieFTT = new LiveCharts.Wpf.LineSeries
            {
                Title = "% Cumplimiento de Producción",
                Values = valoresCumplimiento,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 152, 219)), // Azul claro
                StrokeThickness = 4,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 14,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(41, 128, 185)), // Color del punto (relleno)
                                                                                                                              // ❌ NO USAR PointBackground (no existe en LiveCharts 1.x)
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                LineSmoothness = 0, // Ligeramente suavizada
            };

            // Asignar series (solo una para FTT)
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieFTT
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "PORCENTAJE (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // FTT generalmente va de 0 a 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de FTT Deshidratado - Análisis Semanal
        /// </summary>
        private void GraficarCumplimientoOtrasAreas(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para FTT
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO DE PRODUCCIÓN - OTRAS ÁREAS",
                    "Cumplimiento semanal de producción"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = $@"
        SELECT 
            EXTRACT(YEAR FROM ""Fecha"") as año,
            EXTRACT(WEEK FROM ""Fecha"") as No_Semana,

            -- Porcentaje limitado al 100%
            ROUND(LEAST(
                (SUM(""Kg_prod_term"") - SUM(""Kg_fuera_espec"")) / NULLIF(SUM(""Kg_meta""), 0) * 100,
                100
            ),2) as ""% Cumplimiento""
        FROM public.""Ficha""
        WHERE ""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
            AND EXTRACT(YEAR FROM ""Fecha"") = {añoSeleccionado}
            AND EXTRACT(WEEK FROM ""Fecha"") IN ({semanasParam})
        GROUP BY 
            EXTRACT(YEAR FROM ""Fecha""),
            EXTRACT(WEEK FROM ""Fecha"")
        ORDER BY 
            año,
            No_Semana;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en Cumplimiento",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarCumplimientoOtrasAreas(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para FTT Deshidratado (Línea)
        /// </summary>
        private void ConfigurarCumplimientoOtrasAreas(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Crear valores
            var valoresCumplimiento = new LiveCharts.ChartValues<double>();
            var semanas = new List<string>();

            foreach (DataRow row in datos.Rows)
            {
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);
                valoresCumplimiento.Add(cumplimiento);
                semanas.Add($"Sem {row["No_Semana"]}");
            }

            // Crear serie de línea para FTT - CORREGIDO
            var serieFTT = new LiveCharts.Wpf.LineSeries
            {
                Title = "% Cumplimiento de Producción",
                Values = valoresCumplimiento,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 152, 219)), // Azul claro
                StrokeThickness = 4,
                Fill = System.Windows.Media.Brushes.Transparent,
                PointGeometrySize = 14,
                PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                PointForeground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(41, 128, 185)), // Color del punto (relleno)
                                                                                                                              // ❌ NO USAR PointBackground (no existe en LiveCharts 1.x)
                DataLabels = true,
                FontSize = 10,
                Foreground = System.Windows.Media.Brushes.Black,
                LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                LineSmoothness = 0, // Ligeramente suavizada
            };

            // Asignar series (solo una para FTT)
            cartesianChart1.Series = new LiveCharts.SeriesCollection
    {
        serieFTT
    };

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "PORCENTAJE (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // FTT generalmente va de 0 a 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarCumplimientoPorSupervisor(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO POR SUPERVISOR - ANÁLISIS SEMANAL",
                    "Porcentaje de cumplimiento"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = @"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") as año,
    EXTRACT(WEEK FROM f.""Fecha"") as No_Semana,
    u.""Usuario"" as Supervisor,  -- Ya no usamos COALESCE

    -- Porcentaje limitado al 100%
    ROUND(LEAST(
        (SUM(f.""Kg_prod_seco"") - SUM(f.""Kg_fuera_espec"")) / NULLIF(SUM(f.""Kg_meta""), 0) * 100,
        100
    ), 2) as ""% Cumplimiento""

FROM public.""Ficha"" f
INNER JOIN public.""Usuarios"" u ON f.""ID_user"" = u.""ID_User""  -- Cambiado a INNER JOIN
WHERE f.""Area"" = 'Despegue'
    AND u.""Nivel"" = 'Supervisor'
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(WEEK FROM f.""Fecha""),
    u.""Usuario""
ORDER BY 
    año,
    No_Semana,
    Supervisor;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento por Supervisor.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoPorSupervisor(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento por Supervisor: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoPorSupervisor(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y semanas únicas
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Supervisor"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            var semanas = datos.AsEnumerable()
                               .Select(row => $"Sem {row.Field<decimal>("No_Semana")}")
                               .Distinct()
                               .OrderBy(s => s)
                               .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todas las semanas con 0
                foreach (var semana in semanas)
                {
                    valoresPorSupervisor[supervisor][semana] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Supervisor"].ToString();
                string semana = $"Sem {row["No_Semana"]}";
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][semana] = cumplimiento;
                }
            }

            // Array de colores para los supervisores (colores vibrantes para líneas)
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(52, 152, 219),  // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),  // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),   // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),  // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),  // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),  // Naranja
        System.Windows.Media.Color.FromRgb(26, 188, 156),  // Turquesa
        System.Windows.Media.Color.FromRgb(52, 73, 94),    // Azul oscuro
        System.Windows.Media.Color.FromRgb(192, 57, 43),   // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),  // Azul medio
            };

            // Crear una serie de línea por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var semana in semanas)
                {
                    valores.Add(valoresPorSupervisor[supervisor][semana]);
                }

                var color = colores[colorIndex % colores.Length];

                var serie = new LiveCharts.Wpf.LineSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Stroke = new System.Windows.Media.SolidColorBrush(color),
                    StrokeThickness = 3,
                    Fill = System.Windows.Media.Brushes.Transparent,
                    PointGeometrySize = 12,
                    PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                    PointForeground = System.Windows.Media.Brushes.White,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                    LineSmoothness = 0, // Suavizado moderado
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "CUMPLIMIENTO (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // Cumplimiento limitado al 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // Agregar una línea horizontal en el 100% (opcional - meta ideal)
            // Nota: LiveCharts 1.x no tiene líneas de anotación fáciles, 
            // pero podemos agregar un separador especial en el eje Y

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento Mensual por Supervisor (Despegue)
        /// </summary>
        private void GraficarCumplimientoMensualPorSupervisor()
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento Mensual
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO MENSUAL POR SUPERVISOR - DESPEGUE",
                    "Análisis mensual de cumplimiento"
                );

                // Construir la consulta SQL
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = $@"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") as año,
    INITCAP(
        CASE EXTRACT(MONTH FROM f.""Fecha"")
            WHEN 1 THEN 'enero'
            WHEN 2 THEN 'febrero'
            WHEN 3 THEN 'marzo'
            WHEN 4 THEN 'abril'
            WHEN 5 THEN 'mayo'
            WHEN 6 THEN 'junio'
            WHEN 7 THEN 'julio'
            WHEN 8 THEN 'agosto'
            WHEN 9 THEN 'septiembre'
            WHEN 10 THEN 'octubre'
            WHEN 11 THEN 'noviembre'
            WHEN 12 THEN 'diciembre'
        END
    ) as Mes,
    u.""Usuario"" as Supervisor,  -- Sin COALESCE, solo usuarios reales

    -- Porcentaje limitado al 100%
    ROUND(LEAST(
        (SUM(f.""Kg_prod_seco"") - SUM(f.""Kg_fuera_espec"")) / NULLIF(SUM(f.""Kg_meta""), 0) * 100,
        100
    ), 2) as ""% Cumplimiento""

FROM public.""Ficha"" f
INNER JOIN public.""Usuarios"" u ON f.""ID_user"" = u.""ID_User""  -- Cambiado a INNER JOIN
WHERE f.""Area"" = 'Despegue'
    AND u.""Nivel"" = 'Supervisor'  -- Filtro por nivel Supervisor
    AND EXTRACT(YEAR FROM f.""Fecha"") = {añoSeleccionado}
GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(MONTH FROM f.""Fecha""),
    u.""Usuario""
ORDER BY 
    EXTRACT(MONTH FROM f.""Fecha""),
    Supervisor;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        $"No se encontraron datos para el año {añoSeleccionado} en Cumplimiento Mensual por Supervisor.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoMensual(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar Cumplimiento Mensual: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento Mensual (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoMensual(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y meses únicos
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Supervisor"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            // Ordenar los meses correctamente (enero, febrero, marzo...)
            var ordenMeses = new Dictionary<string, int>
    {
        {"enero", 1}, {"febrero", 2}, {"marzo", 3}, {"abril", 4},
        {"mayo", 5}, {"junio", 6}, {"julio", 7}, {"agosto", 8},
        {"septiembre", 9}, {"octubre", 10}, {"noviembre", 11}, {"diciembre", 12}
    };

            var meses = datos.AsEnumerable()
                             .Select(row => row.Field<string>("Mes").ToLower())
                             .Distinct()
                             .OrderBy(m => ordenMeses[m])
                             .Select(m => char.ToUpper(m[0]) + m.Substring(1)) // Capitalizar primera letra
                             .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todos los meses con 0
                foreach (var mes in meses)
                {
                    valoresPorSupervisor[supervisor][mes] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Supervisor"].ToString();
                string mes = row["Mes"].ToString(); // Ya viene capitalizado por INITCAP
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][mes] = cumplimiento;
                }
            }

            // Array de colores para los supervisores
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(52, 152, 219),  // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),  // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),   // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),  // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),  // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),  // Naranja
        System.Windows.Media.Color.FromRgb(26, 188, 156),  // Turquesa
        System.Windows.Media.Color.FromRgb(52, 73, 94),    // Azul oscuro
        System.Windows.Media.Color.FromRgb(192, 57, 43),   // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),  // Azul medio
            };

            // Crear una serie de línea por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var mes in meses)
                {
                    valores.Add(valoresPorSupervisor[supervisor][mes]);
                }

                var color = colores[colorIndex % colores.Length];

                var serie = new LiveCharts.Wpf.LineSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Stroke = new System.Windows.Media.SolidColorBrush(color),
                    StrokeThickness = 3,
                    Fill = System.Windows.Media.Brushes.Transparent,
                    PointGeometrySize = 12,
                    PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                    PointForeground = System.Windows.Media.Brushes.White,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("F2") + "%",
                    LineSmoothness = 0, // Suavizado suave
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X (Meses)
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "MESES",
                Labels = meses.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "CUMPLIMIENTO (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100,
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento Mensual por Supervisor (Despegue)
        /// </summary>
        private void GraficarCumplimientoMensualPorSupervisorOtrasAreas()
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento Mensual
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO MENSUAL POR SUPERVISOR",
                    "Análisis mensual de cumplimiento(EXCLUYENDO DESPEGUE)"
                );

                // Construir la consulta SQL
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = $@"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") as año,
    INITCAP(
        CASE EXTRACT(MONTH FROM f.""Fecha"")
            WHEN 1 THEN 'enero'
            WHEN 2 THEN 'febrero'
            WHEN 3 THEN 'marzo'
            WHEN 4 THEN 'abril'
            WHEN 5 THEN 'mayo'
            WHEN 6 THEN 'junio'
            WHEN 7 THEN 'julio'
            WHEN 8 THEN 'agosto'
            WHEN 9 THEN 'septiembre'
            WHEN 10 THEN 'octubre'
            WHEN 11 THEN 'noviembre'
            WHEN 12 THEN 'diciembre'
        END
    ) as Mes,
    u.""Usuario"" as Supervisor,  -- Sin COALESCE, solo usuarios reales

    -- Porcentaje limitado al 100%
    ROUND(LEAST(
        (SUM(f.""Kg_prod_term"") - SUM(f.""Kg_fuera_espec"")) / NULLIF(SUM(f.""Kg_meta""), 0) * 100,
        100
    ), 2) as ""% Cumplimiento""

FROM public.""Ficha"" f
INNER JOIN public.""Usuarios"" u ON f.""ID_user"" = u.""ID_User""  -- Cambiado a INNER JOIN
WHERE f.""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
    AND u.""Nivel"" = 'Supervisor'  -- Filtro por nivel Supervisor
    AND EXTRACT(YEAR FROM f.""Fecha"") = {añoSeleccionado}
GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(MONTH FROM f.""Fecha""),
    u.""Usuario""
ORDER BY 
    EXTRACT(MONTH FROM f.""Fecha""),
    Supervisor;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        $"No se encontraron datos para el año {añoSeleccionado} en Cumplimiento Mensual por Supervisor.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoMensualOtrasAreas(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar Cumplimiento Mensual: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento Mensual (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoMensualOtrasAreas(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y meses únicos
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Supervisor"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            // Ordenar los meses correctamente (enero, febrero, marzo...)
            var ordenMeses = new Dictionary<string, int>
    {
        {"enero", 1}, {"febrero", 2}, {"marzo", 3}, {"abril", 4},
        {"mayo", 5}, {"junio", 6}, {"julio", 7}, {"agosto", 8},
        {"septiembre", 9}, {"octubre", 10}, {"noviembre", 11}, {"diciembre", 12}
    };

            var meses = datos.AsEnumerable()
                             .Select(row => row.Field<string>("Mes").ToLower())
                             .Distinct()
                             .OrderBy(m => ordenMeses[m])
                             .Select(m => char.ToUpper(m[0]) + m.Substring(1)) // Capitalizar primera letra
                             .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todos los meses con 0
                foreach (var mes in meses)
                {
                    valoresPorSupervisor[supervisor][mes] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Supervisor"].ToString();
                string mes = row["Mes"].ToString(); // Ya viene capitalizado por INITCAP
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][mes] = cumplimiento;
                }
            }

            // Array de colores para los supervisores
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(52, 152, 219),  // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),  // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),   // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),  // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),  // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),  // Naranja
        System.Windows.Media.Color.FromRgb(26, 188, 156),  // Turquesa
        System.Windows.Media.Color.FromRgb(52, 73, 94),    // Azul oscuro
        System.Windows.Media.Color.FromRgb(192, 57, 43),   // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),  // Azul medio
            };

            // Crear una serie de línea por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var mes in meses)
                {
                    valores.Add(valoresPorSupervisor[supervisor][mes]);
                }

                var color = colores[colorIndex % colores.Length];

                var serie = new LiveCharts.Wpf.LineSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Stroke = new System.Windows.Media.SolidColorBrush(color),
                    StrokeThickness = 3,
                    Fill = System.Windows.Media.Brushes.Transparent,
                    PointGeometrySize = 12,
                    PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                    PointForeground = System.Windows.Media.Brushes.White,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("F2") + "%",
                    LineSmoothness = 0, // Suavizado suave
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X (Meses)
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "MESES",
                Labels = meses.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "CUMPLIMIENTO (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100,
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        /// <summary>
        /// Grafica de % Cumplimiento por Supervisor
        /// </summary>
        private void GraficarCumplimientoPorSupervisorOtrasAreas(List<string> semanasSeleccionadas)
        {
            try
            {
                // Configurar títulos específicos para Cumplimiento
                ConfigurarTitulosGrafico(
                    "% CUMPLIMIENTO POR SUPERVISOR",
                    "Comparativo semanal de cumplimiento por supervisor en áreas de producción(excluyendo Deshidratado)"
                );

                // Construir la consulta SQL
                string semanasParam = string.Join(",", semanasSeleccionadas);
                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                string añoSeleccionado = CB_Anio_grafica.Text;

                string query = @"
SELECT 
    EXTRACT(YEAR FROM f.""Fecha"") as año,
    EXTRACT(WEEK FROM f.""Fecha"") as No_Semana,
    u.""Usuario"" as Supervisor,  -- Sin COALESCE

    -- Porcentaje limitado al 100%
    ROUND(LEAST(
        (SUM(f.""Kg_prod_term"") - SUM(f.""Kg_fuera_espec"")) / NULLIF(SUM(f.""Kg_meta""), 0) * 100,
        100
    ), 2) as ""% Cumplimiento""

FROM public.""Ficha"" f
INNER JOIN public.""Usuarios"" u ON f.""ID_user"" = u.""ID_User""  -- Solo usuarios existentes
WHERE f.""Area"" NOT IN ('Tunel/Sumergidor', 'Despegue')
    AND u.""Nivel"" = 'Supervisor'  -- Solo nivel Supervisor
    AND EXTRACT(YEAR FROM f.""Fecha"") = " + añoSeleccionado + @"
    AND EXTRACT(WEEK FROM f.""Fecha"") IN (" + semanasParam + @")
GROUP BY 
    EXTRACT(YEAR FROM f.""Fecha""),
    EXTRACT(WEEK FROM f.""Fecha""),
    u.""Usuario""
ORDER BY 
    año,
    No_Semana,
    Supervisor;";

                // Ejecutar la consulta
                DataTable datos = dbHelper.ExecuteSelectQuery(query);

                if (datos == null || datos.Rows.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se encontraron datos para las semanas seleccionadas en % Cumplimiento por Supervisor.",
                        "Precaución",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Configurar el gráfico con LiveCharts
                ConfigurarGraficoCumplimientoPorSupervisorOtrasAreas(datos);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al graficar % Cumplimiento por Supervisor: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Configura el gráfico específico para % Cumplimiento por Supervisor (Líneas)
        /// </summary>
        private void ConfigurarGraficoCumplimientoPorSupervisorOtrasAreas(DataTable datos)
        {
            // Limpiar el chart de LiveCharts
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();

            // Obtener supervisores únicos y semanas únicas
            var supervisores = datos.AsEnumerable()
                                    .Select(row => row.Field<string>("Supervisor"))
                                    .Distinct()
                                    .OrderBy(s => s)
                                    .ToList();

            var semanas = datos.AsEnumerable()
                               .Select(row => $"Sem {row.Field<decimal>("No_Semana")}")
                               .Distinct()
                               .OrderBy(s => s)
                               .ToList();

            // Crear un diccionario para acceso rápido a los valores
            var valoresPorSupervisor = new Dictionary<string, Dictionary<string, double>>();

            foreach (var supervisor in supervisores)
            {
                valoresPorSupervisor[supervisor] = new Dictionary<string, double>();

                // Inicializar todas las semanas con 0
                foreach (var semana in semanas)
                {
                    valoresPorSupervisor[supervisor][semana] = 0;
                }
            }

            // Llenar los valores reales
            foreach (DataRow row in datos.Rows)
            {
                string supervisor = row["Supervisor"].ToString();
                string semana = $"Sem {row["No_Semana"]}";
                double cumplimiento = System.Convert.ToDouble(row["% Cumplimiento"]);

                if (valoresPorSupervisor.ContainsKey(supervisor))
                {
                    valoresPorSupervisor[supervisor][semana] = cumplimiento;
                }
            }

            // Array de colores para los supervisores (colores vibrantes para líneas)
            var colores = new System.Windows.Media.Color[]
            {
        System.Windows.Media.Color.FromRgb(52, 152, 219),  // Azul
        System.Windows.Media.Color.FromRgb(46, 204, 113),  // Verde
        System.Windows.Media.Color.FromRgb(231, 76, 60),   // Rojo
        System.Windows.Media.Color.FromRgb(155, 89, 182),  // Púrpura
        System.Windows.Media.Color.FromRgb(241, 196, 15),  // Amarillo
        System.Windows.Media.Color.FromRgb(230, 126, 34),  // Naranja
        System.Windows.Media.Color.FromRgb(26, 188, 156),  // Turquesa
        System.Windows.Media.Color.FromRgb(52, 73, 94),    // Azul oscuro
        System.Windows.Media.Color.FromRgb(192, 57, 43),   // Rojo oscuro
        System.Windows.Media.Color.FromRgb(41, 128, 185),  // Azul medio
            };

            // Crear una serie de línea por cada supervisor
            int colorIndex = 0;
            foreach (var supervisor in supervisores)
            {
                var valores = new LiveCharts.ChartValues<double>();

                foreach (var semana in semanas)
                {
                    valores.Add(valoresPorSupervisor[supervisor][semana]);
                }

                var color = colores[colorIndex % colores.Length];

                var serie = new LiveCharts.Wpf.LineSeries
                {
                    Title = supervisor,
                    Values = valores,
                    Stroke = new System.Windows.Media.SolidColorBrush(color),
                    StrokeThickness = 3,
                    Fill = System.Windows.Media.Brushes.Transparent,
                    PointGeometrySize = 12,
                    PointGeometry = LiveCharts.Wpf.DefaultGeometries.Circle,
                    PointForeground = System.Windows.Media.Brushes.White,
                    DataLabels = true,
                    FontSize = 9,
                    Foreground = System.Windows.Media.Brushes.Black,
                    LabelPoint = point => point.Y.ToString("F2") + "%", // Formato con 2 decimales y %
                    LineSmoothness = 0, // Suavizado moderado
                };

                cartesianChart1.Series.Add(serie);
                colorIndex++;
            }

            // Configurar leyenda
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            // Eje X
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Title = "SEMANAS",
                Labels = semanas.ToArray(),
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                Separator = new LiveCharts.Wpf.Separator { IsEnabled = false }
            });

            // Eje Y (para porcentajes)
            cartesianChart1.AxisY.Add(new LiveCharts.Wpf.Axis
            {
                Title = "CUMPLIMIENTO (%)",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(52, 73, 94)),
                LabelFormatter = value => value.ToString("F2") + "%",
                MinValue = 0,
                MaxValue = 100, // Cumplimiento limitado al 100%
                Separator = new LiveCharts.Wpf.Separator
                {
                    Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 220, 220)),
                    StrokeThickness = 1,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 }
                }
            });

            // Agregar una línea horizontal en el 100% (opcional - meta ideal)
            // Nota: LiveCharts 1.x no tiene líneas de anotación fáciles, 
            // pero podemos agregar un separador especial en el eje Y

            // ✅ HABILITAR TOOLTIP (el cuadro que aparece al pasar el mouse)
            cartesianChart1.DataTooltip = new LiveCharts.Wpf.DefaultTooltip
            {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50)), // Fondo oscuro
                Foreground = System.Windows.Media.Brushes.White, // Texto blanco
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Bold,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                BorderThickness = new System.Windows.Thickness(1),
                CornerRadius = new System.Windows.CornerRadius(5), // Bordes redondeados
                ShowSeries = true, // Mostrar el nombre de la serie
            };

            // Configuración general
            cartesianChart1.Background = System.Windows.Media.Brushes.White;
            cartesianChart1.DisableAnimations = false;
            cartesianChart1.AnimationsSpeed = TimeSpan.FromMilliseconds(800);
            //cartesianChart1.DataTooltip = null;
        }
        // Método para configurar título con Labels externos (opcional)
        private void ConfigurarTitulosGrafico(string titulo, string subtitulo)
        {
            labelSubtitulo.Text = subtitulo;
            labelSubtitulo.Font = new Font("Segoe UI", 12, FontStyle.Italic);
            labelSubtitulo.ForeColor = Color.FromArgb(127, 140, 141);
            labelSubtitulo.TextAlign = ContentAlignment.MiddleCenter;
            labelSubtitulo.Dock = DockStyle.Top;

            labelTitulo.Text = titulo;
            labelTitulo.Font = new Font("Segoe UI", 18, FontStyle.Bold);
            labelTitulo.ForeColor = Color.FromArgb(44, 62, 80);
            labelTitulo.TextAlign = ContentAlignment.MiddleCenter;
            labelTitulo.Dock = DockStyle.Top;
        }

        private void btn_new_report_merma_T_Click(object sender, EventArgs e)
        {
            actualiza_reporte_merma_S();
            btn_clean_merma_S.Enabled = true;
            btn_export_excel_merma_S.Enabled = true;
        }
        //dgv_reporte_merma_S
        private async void btn_export_excel_merma_T_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarDataGridViewAExcel_ClosedXML(
                            dgv_reporte_merma_S,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }

        private void btn_clean_merma_T_Click(object sender, EventArgs e)
        {
            dgv_reporte_merma_S.DataSource = null;   // Desvincula cualquier origen de datos
            dgv_reporte_merma_S.Rows.Clear();        // Limpia todas las filas (por si no hay DataSource)
            dgv_reporte_merma_S.Columns.Clear();     // Limpia todas las columnas
            btn_export_excel_merma_S.Enabled = false;
            btn_clean_merma_S.Enabled = false;
        }
        private void conf_email(string cuerpoHtml) 
        {
            try
            {
                MailMessage correo = new MailMessage();
                correo.From = new MailAddress(RemitenteEMail);
                string destinatarios = DestinatariosEmail;

                foreach (var mail in destinatarios.Split(','))
                {
                    correo.To.Add(mail.Trim());
                }
                correo.Subject = "Reporte Automático del Sistema Tablero";
                correo.Body = cuerpoHtml;
                correo.IsBodyHtml = true;


                // Configuración del servidor SMTP de Office365
                SmtpClient smtp = new SmtpClient(servidor_smtp, PuertoSMTP);
                smtp.Credentials = new NetworkCredential(
                    RemitenteEMail,
                    PasswordEmail
                );
                smtp.EnableSsl = true;  // Office365 requiere STARTTLS
                smtp.Send(correo);
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, "Error: " + ex.Message,
                                                            "Error al enviar el correo electrónico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public string GenerarTablaValores(List<KeyValuePair<string, string>> valores)
        {
            string tabla = @"
            <table style='border-collapse: collapse; width: 100%; margin-top: 10px;'>
            <tr style='background: #e9ecef;'>
                <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Concepto</th>
                <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Valor</th>
            </tr>";

            foreach (var item in valores)
            {
                tabla += $@"
            <tr>
                <td style='border: 1px solid #ccc; padding: 8px;'>{item.Key}</td>
                <td style='border: 1px solid #ccc; padding: 8px;'>{item.Value}</td>
            </tr>";
            }

            tabla += "</table>";

            return tabla;
        }
        public string GenerarTablaResumen(List<KeyValuePair<string, string>> valores)
        {
            string tabla = @"
            <table style='border-collapse: collapse; width: 100%; margin-top: 10px;'>
            <tr style='background: #e9ecef;'>
                <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Concepto</th>
                <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Valor</th>
            </tr>";

            foreach (var item in valores)
            {
                tabla += $@"
            <tr>
                <td style='border: 1px solid #ccc; padding: 8px;'>{item.Key}</td>
                <td style='border: 1px solid #ccc; padding: 8px;'>{item.Value}</td>
            </tr>";
            }
            if (cb_Area.SelectedIndex == 1)
            {
                decimal porcentaje = 0;
                string valorLimpio = Txt_Read_5.Text.Replace("%", "").Trim();
                decimal.TryParse(valorLimpio, out porcentaje);

                string colorFila;

                if (porcentaje >= 90)
                    colorFila = "#00913F"; // Verde
                else if (porcentaje >= 80)
                    colorFila = "#E5BE01"; // Amarillo
                else
                    colorFila = "#A52019"; // Rojo

                tabla += $@"
                    <tr style='background-color: {colorFila};'>
                        <td style='border: 1px solid #ccc; padding: 8px;'>% Cumplimiento</td>
                        <td style='border: 1px solid #ccc; padding: 8px;'>{Txt_Read_5.Text}</td>
                    </tr>";
            }
            if (cb_Area.SelectedIndex != 0 && cb_Area.SelectedIndex != 1)
            {
                decimal porcentaje = 0;
                string valorLimpio = Txt_Read_4.Text.Replace("%", "").Trim();
                decimal.TryParse(valorLimpio, out porcentaje);

                string colorFila;

                if (porcentaje >= 90)
                    colorFila = "#d4edda"; // Verde
                else if (porcentaje >= 80)
                    colorFila = "#fff3cd"; // Amarillo
                else
                    colorFila = "#f8d7da"; // Rojo

                tabla += $@"
                    <tr style='background-color: {colorFila};'>
                        <td style='border: 1px solid #ccc; padding: 8px;'>% Cumplimiento</td>
                        <td style='border: 1px solid #ccc; padding: 8px;'>{Txt_Read_4.Text}</td>
                    </tr>";
            }
            tabla += "</table>";

            return tabla;
        }
        public string GenerarTablaTiemposMuertosMecanicos()
        {
            // Verifica si la lista está vacía o es nula
            if (valores_mecanico == null || valores_mecanico.Count == 0)
                return "<p>No se registraron tiempos muertos mecánicos.</p>";

            string tabla = @"
<div style='margin-top: 20px;'>
    <h3 style='color: #2c3e50;'>Tiempos Muertos Mecánicos</h3>
    <table style='border-collapse: collapse; width: 100%; margin-top: 10px;'>
        <tr style='background: #e9ecef;'>
            <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Minutos Detenidos</th>
            <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Tipo</th>
            <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Motivos</th>
        </tr>";

            foreach (var item in valores_mecanico)
            {
                tabla += $@"
        <tr>
            <td style='border: 1px solid #ccc; padding: 8px;'>{item.Item1}</td>
            <td style='border: 1px solid #ccc; padding: 8px;'>{item.Item2}</td>
            <td style='border: 1px solid #ccc; padding: 8px;'>{item.Item3}</td>
        </tr>";
            }

            tabla += "</table></div>";
            tabla += $@"<br><p> Con un total de minutos detenidos de <strong>{txt_TM_mecanico.Text}</strong> minutos.</p>";

            return tabla;
        }
        public string GenerarTablaTiemposOperativos()
        {
            // Verifica si la lista está vacía o es nula
            if (valores_operativos == null || valores_operativos.Count == 0)
                return "<p>No se registraron tiempos muertos operativos.</p>";

            string tabla = @"
<div style='margin-top: 20px;'>
    <h3 style='color: #2c3e50;'>Tiempos Muertos Operativos</h3>
    <table style='border-collapse: collapse; width: 100%; margin-top: 10px;'>
        <tr style='background: #e9ecef;'>
            <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Minutos Detenidos</th>
            <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Tipo</th>
            <th style='border: 1px solid #ccc; padding: 8px; text-align: left;'>Motivos</th>
        </tr>";

            foreach (var item in valores_operativos)
            {
                tabla += $@"
        <tr>
            <td style='border: 1px solid #ccc; padding: 8px;'>{item.Item1}</td>
            <td style='border: 1px solid #ccc; padding: 8px;'>{item.Item2}</td>
            <td style='border: 1px solid #ccc; padding: 8px;'>{item.Item3}</td>
        </tr>";
            }

            tabla += "</table></div>";
            tabla += $@"<br><p> Con un total de minutos detenidos de <strong>{txt_TM_operativo.Text}</strong> minutos.</p>";

            return tabla;
        }
        public string GenerarTablaDetallesOP()
        {
            string valorBuscado = cb_OP.Text;
            string Orden_Produccion = string.Empty;
            string Producto = string.Empty;
            string Medida = string.Empty;
            string Descripcion = string.Empty;
            string Especificacion = string.Empty;
            string Ingredientes = string.Empty;
            string Humedad = string.Empty;
            string Comercio = string.Empty;
            string Manzana = string.Empty;
            string Analisis = string.Empty;
            string Area_Proceso = string.Empty;
            string OP_Origen = string.Empty;
            string Destino = string.Empty;
            string Clasificacion = string.Empty;

            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            // Consulta para buscar donde OP = valor_buscado
            string query = "SELECT \"Orden_Produccion\", \"Producto\", \"Medida\", \"Descripcion\", " +
                "\"Especificacion\", \"Ingredientes\", \"Humedad\", \"Comercio\", \"Manzana\", \"Analisis\", " +
                "\"Area_Proceso\", \"OP_Origen\", \"Destino\", \"Clasificacion\" \r\nFROM public.\"Detalles_OP\" " +
                "where \"Orden_Produccion\" = @valorBuscado";

            // Crear parámetro
            NpgsqlParameter[] parameters = new NpgsqlParameter[]
            {
              new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = valorBuscado }
            };

            // Ejecutar consulta
            System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

            // Verificar si se encontraron resultados
            if (dt != null && dt.Rows.Count > 0)
            {
                Orden_Produccion = dt.Rows[0]["Orden_Produccion"].ToString();
                Producto = dt.Rows[0]["Producto"].ToString();
                Medida = dt.Rows[0]["Medida"].ToString();
                Descripcion = dt.Rows[0]["Descripcion"].ToString();
                Especificacion = dt.Rows[0]["Especificacion"].ToString();
                Ingredientes = dt.Rows[0]["Ingredientes"].ToString();
                Humedad = dt.Rows[0]["Humedad"].ToString();
                Comercio = dt.Rows[0]["Comercio"].ToString();
                Manzana = dt.Rows[0]["Manzana"].ToString();
                Analisis = dt.Rows[0]["Analisis"].ToString();
                Area_Proceso = dt.Rows[0]["Area_Proceso"].ToString();
                OP_Origen = dt.Rows[0]["OP_Origen"].ToString();
                Destino = dt.Rows[0]["Destino"].ToString();
                Clasificacion = dt.Rows[0]["Clasificacion"].ToString();
            }
            else
            {
                return "<p>No se encontró registro en la tabla de detalles OP, favor de contactar al usuario Administrador y registrar los detalles del OP.</p>";
            }


            string datos = $@"<p>Orden de Produccion: <strong>{Orden_Produccion}</strong>, Producto: <strong>{Producto}</strong>, Medida: <strong>{Medida}</strong>, Descripción: <strong>{Descripcion}</strong>, Especificación: <strong>{Especificacion}</strong>, Ingredientes: <strong>{Ingredientes}</strong>, Humedad: <strong>{Humedad}</strong>, Comercio: <strong>{Comercio}</strong>, Manzana: <strong>{Manzana}</strong>, Analisis: <strong>{Analisis}</strong>, Area de Proceso: <strong>{Area_Proceso}</strong>, OP de Origen: <strong>{OP_Origen}</strong>, Destino: <strong>{Destino}</strong>, Clasificación: <strong>{Clasificacion}</strong>,</p>";

            return datos;
        }
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            // Alternar entre image1 e image3 con cada clic
            showingImage3 = !showingImage3;

            // Reiniciar la animación
            transitionProgress = 0f;

            // Forzar una actualización inmediata
            if (showingImage3 || !(mostrarPassword))
            {
                pictureBox4.Image = image3;
                // Mostrar contraseña
                Txt_Password_SMTP.PasswordChar = '\0'; // Esto quita el carácter de ocultamiento
                mostrarPassword = true;
            }
            else
            {
                pictureBox4.Image = image1;
                // Ocultar contraseña
                Txt_Password_SMTP.PasswordChar = '*';
                mostrarPassword = false;
            }
        }

        private void pictureBox4_MouseEnter(object sender, EventArgs e)
        {
            isHovering = true;
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            isHovering = false;
        }

        private void Txt_Destinatarios_Enter(object sender, EventArgs e)
        {
            lbl_text_asistive.ForeColor = Color.FromArgb(59, 140, 255);
            lbl_text_asistive.Visible = true;
        }

        private void Txt_Destinatarios_Leave(object sender, EventArgs e)
        {
            lbl_text_asistive.Visible = true;
        }

        private void Txt_Remitente_TextChanged(object sender, EventArgs e)
        {
            ValidarCorreoRemitente();
        }
        private bool ValidarCorreoRemitente()
        {
            string email = Txt_Remitente.Text.Trim();

            // Regex simple y confiable para validar correos
            string pattern = @"^[^@\s]+@[^@\s]+\.[^@\s]+$";

            if (Regex.IsMatch(email, pattern))
            {
                errorProvider1.SetError(Txt_Remitente, ""); // Sin error
                return true;
            }
            else
            {
                errorProvider1.SetError(Txt_Remitente, "Ingrese un único correo válido.");
                return false;
            }
        }

        private void btn_save_settings_email_Click(object sender, EventArgs e)
        {
            if (!ValidarCorreoRemitente())
            {
                return;
            }
            ValidarDestinatarios();
            if (!string.IsNullOrEmpty(errorProvider1.GetError(Txt_Destinatarios)))
            {
                return;
            }
            if(string.IsNullOrEmpty(Txt_server_Smtp.Text)||string.IsNullOrEmpty(Txt_Password_SMTP.Text)||string.IsNullOrEmpty(TxtPuerto.Text))
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Tablero.Properties.Settings.Default.ServerSMTP = Txt_server_Smtp.Text;
            Tablero.Properties.Settings.Default.RemitenteEMail = Txt_Remitente.Text;
            Tablero.Properties.Settings.Default.PasswordEmail = Txt_Password_SMTP.Text;
            Tablero.Properties.Settings.Default.PuertoSMTP = System.Convert.ToInt32(TxtPuerto.Text);
            Tablero.Properties.Settings.Default.SSLCheck = Check_ssl.Checked;
            Tablero.Properties.Settings.Default.DestinatariosEmail = Txt_Destinatarios.Text;
            Tablero.Properties.Settings.Default.Save();
            email_varibles();
            test_correo();
        }
        private void Txt_Destinatarios_TextChanged(object sender, EventArgs e)
        {
            ValidarDestinatarios();
        }
        private void ValidarDestinatarios()
        {
            string input = Txt_Destinatarios.Text.Trim();

            // Si está vacío, mostrar error
            if (string.IsNullOrWhiteSpace(input))
            {
                errorProvider1.SetError(Txt_Destinatarios, "Ingrese uno o varios correos.");
                return;
            }

            // Separar por comas
            string[] correos = input.Split(',');

            // Expresión regular para validar correo electrónico
            string patron = @"^[^@\s]+@[^@\s]+\.[^@\s]+$";

            foreach (string correo in correos)
            {
                string c = correo.Trim();

                // Validación individual
                if (!System.Text.RegularExpressions.Regex.IsMatch(c, patron))
                {
                    errorProvider1.SetError(Txt_Destinatarios, "Formato de correo inválido en la lista.");
                    return;
                }
            }

            // Si todo está bien, limpiar error
            errorProvider1.SetError(Txt_Destinatarios, "");
        }

        private void TxtPuerto_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permitir solo números y la tecla Backspace
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Bloquea la tecla
            }
        }

        private void btn_prueba_Click(object sender, EventArgs e)
        {
            //servidor_smtp
            //RemitenteEMail
            //PasswordEmail
            //PuertoSMTP
            //SSLCheck
            //DestinatariosEmail

            // Validar que la configuración esté completa
            if (string.IsNullOrWhiteSpace(servidor_smtp) ||
                string.IsNullOrWhiteSpace(RemitenteEMail) ||
                string.IsNullOrWhiteSpace(PasswordEmail) ||
                PuertoSMTP == 0)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "SMTP sin configurar.",
                    "SMTP pendiente",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }
            test_correo();
        }
        private void test_correo()
        {
            // Enviar correo de prueba
            try
            {
                string cuerpoHtml = $@"
        <html>
        <body style='font-family: Arial; font-size: 14px; color: #333;'>
            <p>Este es un correo de prueba enviado desde el sistema <b>Tablero</b>.</p>
            <p>La conexión SMTP se ha verificado correctamente.</p>
            <br>
            <p style='color: #888;'>Mensaje generado automáticamente, favor de no responder.</p>
        </body>
        </html>";


                // Configuración del correo
                MailMessage correo = new MailMessage();
                correo.From = new MailAddress(RemitenteEMail);
                string destinatarios = DestinatariosEmail;

                foreach (var mail in destinatarios.Split(','))
                {
                    correo.To.Add(mail.Trim());
                }
                correo.Subject = "Prueba de conexión SMTP - Sistema Tablero";
                correo.Body = cuerpoHtml;
                correo.IsBodyHtml = true;


                // Configuración del servidor SMTP de Office365
                SmtpClient smtp = new SmtpClient(servidor_smtp, PuertoSMTP);
                smtp.Credentials = new NetworkCredential(
                    RemitenteEMail,
                    PasswordEmail
                );
                smtp.EnableSsl = SSLCheck;  // Office365 requiere STARTTLS

                smtp.Send(correo);

                MetroFramework.MetroMessageBox.Show(this,
                    "El correo de prueba se envió correctamente.",
                    "Prueba exitosa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Error al enviar el correo de prueba:\n" + ex.Message + ".\n favor de espear 24 hrs para la sinconización con su correo ó Revisar con el administrador de correo electronico.",
                    "Error SMTP",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            // Fin del envío de correo de prueba
        }

        private void borrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            editar = false;
            borrar = true;
            Editar Editar_ficha = new Editar(connectionString, editar, borrar);
            // Suscripción al evento con los dos parámetros
            Editar_ficha.FichaSeleccionada += (id_global, area) =>
            {
            };

            Editar_ficha.Show();
        }

        private void btn_new_report_Tiempos_Click(object sender, EventArgs e)
        {
            reporte_tiempos();
            btn_export_excel_Tiempos.Enabled = true;
            btn_clean_Tiempos.Enabled = true;
            txt_filtro_report_Tiempos.Enabled = true;
            tapcontrol.Enabled = true;
            btn_filtro_Tiempos.Enabled = true;
        }
        public void reporte_tiempos()
        {
            string var1 = cb_area_reporte_Tiempos.Text;
            string querySimple = string.Empty;
            rgv_reporte_Tiempos.DataSource = null;
            rgv_reporte_Tiempos.Rows.Clear();
            rgv_reporte_Tiempos.Columns.Clear();
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Obtener el año seleccionado del ComboBox
            string añoSeleccionado = CB_Anio_Tiempos.Text;

            // Validar que se haya seleccionado un año
            if (string.IsNullOrEmpty(añoSeleccionado))
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "Por favor, seleccione un año antes de generar el reporte de tiempos.",
                    "Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (var1 == "Todos")
            {
                // Consulta para TODAS las áreas - ACTUALIZADA con filtro de año
                querySimple = @"SELECT 
    f.""OP"" AS ""OP"",
    f.""Fecha"" AS ""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    COALESCE(tmo.""Tipo"", 'Operativo') AS ""Tipo de Tiempo Muerto"",
    (tmo.""Min_Detenidos"") AS ""Minutos Detenidos"",
    tmo.""Motivos"" AS ""Motivos"",
    f.""Area"" AS ""Area""
FROM ""Ficha"" f
LEFT JOIN ""Tiempo_Muerto_Operativo"" tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""
WHERE tmo.""Min_Detenidos"" IS NOT NULL
    AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio

UNION ALL

SELECT 
    f.""OP"" AS ""OP"",
    f.""Fecha"" AS ""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    COALESCE(tmm.""Tipo"", 'Mecánico') AS ""Tipo de Tiempo Muerto"",
    (tmm.""Min_Detenidos"") AS ""Minutos Detenidos"",
    tmm.""Motivos"" AS ""Motivos"",
    f.""Area"" AS ""Area""
FROM ""Ficha"" f
LEFT JOIN ""Tiempo_muerto_Mecanico"" tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""
WHERE tmm.""Min_Detenidos"" IS NOT NULL
    AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio

ORDER BY ""Fecha"" DESC, ""Area"", ""OP"", ""Tipo de Tiempo Muerto"";";

                // Crear parámetro para el año
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@Anio", NpgsqlTypes.NpgsqlDbType.Integer)
            {
                Value = Convert.ToInt32(añoSeleccionado)
            }
                };
                dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, rgv_reporte_Tiempos, parameters);
            }
            else
            {
                // Consulta para un área específica - ACTUALIZADA con filtro de año
                querySimple = @"SELECT 
    f.""OP"" AS ""OP"",
    f.""Fecha"" AS ""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    COALESCE(tmo.""Tipo"", 'Operativo') AS ""Tipo de Tiempo Muerto"",
    (tmo.""Min_Detenidos"") AS ""Minutos Detenidos"",
    tmo.""Motivos"" AS ""Motivos"",
    f.""Area"" AS ""Area""
FROM ""Ficha"" f
LEFT JOIN ""Tiempo_Muerto_Operativo"" tmo ON f.""ID_Ficha"" = tmo.""ID_Ficha""
WHERE (f.""Area"" = @Area)
   AND tmo.""Min_Detenidos"" IS NOT NULL
   AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio

UNION ALL

SELECT 
    f.""OP"" AS ""OP"",
    f.""Fecha"" AS ""Fecha"",
    EXTRACT(WEEK FROM f.""Fecha"") AS ""No. Semana"",
    COALESCE(tmm.""Tipo"", 'Mecánico') AS ""Tipo de Tiempo Muerto"",
    (tmm.""Min_Detenidos"") AS ""Minutos Detenidos"",
    tmm.""Motivos"" AS ""Motivos"",
    f.""Area"" AS ""Area""
FROM ""Ficha"" f
LEFT JOIN ""Tiempo_muerto_Mecanico"" tmm ON f.""ID_Ficha"" = tmm.""ID_Ficha""
WHERE (f.""Area"" = @Area)
   AND tmm.""Min_Detenidos"" IS NOT NULL
   AND EXTRACT(YEAR FROM f.""Fecha"") = @Anio

ORDER BY ""Fecha"" DESC, ""OP"", ""Tipo de Tiempo Muerto"";";

                // Crear parámetros para área y año
                NpgsqlParameter[] parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
            new NpgsqlParameter("@Anio", NpgsqlTypes.NpgsqlDbType.Integer)
            {
                Value = Convert.ToInt32(añoSeleccionado)
            }
                };
                dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, rgv_reporte_Tiempos, parameters);
            }

            // Formato solo fecha (verificar que la columna existe)
            if (rgv_reporte_Tiempos.Columns.Contains("Fecha"))
            {
                rgv_reporte_Tiempos.Columns["Fecha"].FormatString = "{0:dd/MM/yyyy}";
                rgv_reporte_Tiempos.Columns["Fecha"].FormatInfo = new System.Globalization.CultureInfo("es-MX");
            }
        }
        private void cb_area_reporte_Tiempos_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_area_reporte_Tiempos.SelectedIndex != -1)
            {
                btn_new_report_Tiempos.Enabled = true;
            }
        }

        private async void btn_export_excel_Tiempos_Click(object sender, EventArgs e)
        {
            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarRadGridViewFiltradoAExcel_ClosedXML(
                            rgv_reporte_Tiempos,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }

        private void btn_clean_Tiempos_Click(object sender, EventArgs e)
        {
            // Limpiar el DataGridView
            rgv_reporte_Tiempos.DataSource = null;
            rgv_reporte_Tiempos.Rows.Clear();
            rgv_reporte_Tiempos.Columns.Clear();
            // Limpiar el filtro al cargar nuevos datos
            txt_filtro_report_Tiempos.Clear();
            // Deshabilitar botones hasta que se seleccione un reporte y área
            btn_export_excel_Tiempos.Enabled = false;
            btn_clean_Tiempos.Enabled = false;
            txt_filtro_report_Tiempos.Enabled = false;
            btn_filtro_Tiempos.Enabled = false;
            btn_new_report_Tiempos.Enabled = false;
            cb_area_reporte_Tiempos.SelectedIndex = -1;
        }

        private void btn_filtro_Tiempos_Click(object sender, EventArgs e)
        {
            txt_filtro_report_Tiempos.Clear();
        }

        private void txt_filtro_report_Tiempos_TextChanged(object sender, EventArgs e)
        {
            // Refrescar el grid para aplicar el filtro
            if (rgv_reporte_Tiempos != null)
            {
                rgv_reporte_Tiempos.MasterTemplate.Refresh();
            }
        }

        private void rgv_reporte_Tiempos_CustomFiltering(object sender, GridViewCustomFilteringEventArgs e)
        {
            string textoFiltro = txt_filtro_report_Tiempos.Text.Trim();

            // Si no hay texto de filtro, mostrar todas las filas
            if (string.IsNullOrEmpty(textoFiltro))
            {
                e.Visible = true;
                ResetearEstiloCeldas(e);
                return;
            }

            // Iniciar actualización para mejor rendimiento
            rgv_reporte_Tiempos.BeginUpdate();

            // Por defecto ocultar la fila
            e.Visible = false;

            // Buscar en todas las celdas de la fila
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                GridViewCellInfo celda = e.Row.Cells[i];

                // Verificar si la celda tiene valor
                if (celda.Value != null)
                {
                    string textoCelda = celda.Value.ToString();

                    // Buscar coincidencia (case-insensitive)
                    if (textoCelda.IndexOf(textoFiltro, 0, StringComparison.InvariantCultureIgnoreCase) >= 0)
                    {
                        e.Visible = true; // Mostrar la fila si hay coincidencia

                        // Resaltar la celda que coincide
                        celda.Style.CustomizeFill = true;
                        celda.Style.DrawFill = true;
                        celda.Style.BackColor = Color.FromArgb(201, 252, 254); // Color azul claro
                    }
                    else
                    {
                        // Resetear estilo si no coincide
                        celda.Style.Reset();
                    }
                }
                else
                {
                    // Resetear estilo si la celda es nula
                    celda.Style.Reset();
                }
            }

            rgv_reporte_Tiempos.EndUpdate(false);
        }

        private void dgv_operativo_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv_operativo.IsCurrentCellDirty)
            {
                dgv_operativo.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void cb_supervisor_turno_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cb_supervisor_turno.SelectedItem != null)
            {
                // OPCIÓN 1: Usando SelectedValue (recomendado)
                string valorNoEmpleado = cb_supervisor_turno.SelectedValue?.ToString();
                lbl_no_emp2.Text = valorNoEmpleado;
                lbl_nom2.Text = cb_supervisor_turno.Text;


                DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
                // Consulta para buscar donde OP = valor_buscado
                string query2 = "SELECT \"ID_User\" FROM public.\"Usuarios\" where \"No_Empleado\" = @valorBuscado";

                // Crear parámetro
                NpgsqlParameter[] parameters2 = new NpgsqlParameter[]
                {
                    new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Varchar)
                        {
                            Value = valorNoEmpleado
                        }
                };

                // Ejecutar consulta
                System.Data.DataTable dt2 = dbHelper.ExecuteSelectQuery(query2, parameters2);
                string id_supervisor = string.Empty;
                // Verificar si se encontraron resultados
                if (dt2 != null && dt2.Rows.Count > 0)
                {
                    id_supervisor = dt2.Rows[0]["ID_User"].ToString();
                }
                id_supervisor_global = System.Convert.ToInt32(id_supervisor);
            }
        }

        private void dgv_operativo_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

            // Marcar el error como manejado para que no muestre el diálogo predeterminado
            e.ThrowException = false;

            // Opcional: Si el error es por valor inválido en ComboBox, limpiar la celda
            if (e.Exception is ArgumentException)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    dgv_operativo.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = null;
                }
            }
        }

        private void cb_tipo_grafica_SelectedIndexChanged(object sender, EventArgs e)
        {
            HashSet<int> indicesConSemanas = new HashSet<int>
            {
                0, 1, 2, 3, 4, 5, 6, 7, 9, 11, 12, 13, 14, 15
            };

            lista_semanas.Visible = indicesConSemanas.Contains(cb_tipo_grafica.SelectedIndex);
        }

        private void radMultiColumnComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (radMultiColumnComboBox1.SelectedItem != null)
            {
                // Usando ValueMember que configuramos como "ID_Ficha"
                var selectedValue = radMultiColumnComboBox1.SelectedValue;

                if (selectedValue != null)
                {
                    // También puedes obtener el texto mostrado
                    //string textoMostrado = radMultiColumnComboBox1.Text;
                    //MessageBox.Show($"Texto mostrado: {textoMostrado}", "Display", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    int valorBuscado = System.Convert.ToInt32(selectedValue);


                    DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

                    // Consulta para buscar donde OP = valor_buscado
                    string query = "SELECT \"OP\", \"kg_frescos_enter_se\" FROM public.\"Ficha\" WHERE \"ID_Ficha\" = @valorBuscado and \"Area\" = 'Tunel/Sumergidor';";

                    // Crear parámetro
                    NpgsqlParameter[] parameters = new NpgsqlParameter[]
                    {
                            new NpgsqlParameter("@valorBuscado", NpgsqlTypes.NpgsqlDbType.Integer){Value = valorBuscado}
                    };

                    // Ejecutar consulta
                    System.Data.DataTable dt = dbHelper.ExecuteSelectQuery(query, parameters);

                    // Variable string donde guardar el resultado
                    string OP = string.Empty;
                    string kgFrescoEnterSecador = string.Empty;

                    // Verificar si se encontraron resultados
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        OP = dt.Rows[0]["OP"].ToString();
                        kgFrescoEnterSecador = dt.Rows[0]["kg_frescos_enter_se"].ToString();
                        cb_OP.Enabled = true;
                        cb_OP.Text = OP;
                        cb_OP.Enabled = false;

                        cb_Turno.Enabled = true;

                        Txt_Read_4.Text = kgFrescoEnterSecador;
                        buscar_Meta_hr();
                    }
                }
            }
        }

        private void dgv_reporte_concentrado_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            string colName = dgv_reporte_concentrado.Columns[e.ColumnIndex].Name;

            if (colName != "%Cumplimiento Tiempo Efectivo" && colName != "%Cumplimiento Fresco" && colName != "%Cumplimiento Secos" && colName != "%Cumplimiento Relación Fresco-Seco" && colName != "FTT" && colName != "%Cumplimiento Resecado")
                return;

            if (e.Value == null || e.Value == DBNull.Value)
                return;

            if (decimal.TryParse(e.Value.ToString(), out decimal valor))
            {
                if (valor >= 90)
                {
                    e.CellStyle.BackColor = Color.FromArgb(76, 175, 80);   // Verde
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (valor >= 80)
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 235, 59);  // Amarillo
                    e.CellStyle.ForeColor = Color.Black;
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(244, 67, 54);   // Rojo
                    e.CellStyle.ForeColor = Color.White;
                }
            }
        }

        private void dgv_reporte_concentrado_otras_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            string colName = dgv_reporte_concentrado_otras.Columns[e.ColumnIndex].Name;

            if (colName != "% Cumplimiento Tiempo Efectivo" && colName != "% Cumplimiento Kg Terminados" && colName != "FTT")
                return;

            if (e.Value == null || e.Value == DBNull.Value)
                return;

            if (decimal.TryParse(e.Value.ToString(), out decimal valor))
            {
                if (valor >= 90)
                {
                    e.CellStyle.BackColor = Color.FromArgb(76, 175, 80);   // Verde
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (valor >= 80)
                {
                    e.CellStyle.BackColor = Color.FromArgb(255, 235, 59);  // Amarillo
                    e.CellStyle.ForeColor = Color.Black;
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(244, 67, 54);   // Rojo
                    e.CellStyle.ForeColor = Color.White;
                }
            }
        }

        private void bnt_limpiar_check_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.Items)
            {
                item.Selected = false;
                item.Checked = false;
            }
        }

        private void materialButton8_Click(object sender, EventArgs e)
        {
            try
            {
                List<int> semanasSeleccionadas = ObtenerSemanasSeleccionadasConCheckbox();

                if (semanasSeleccionadas.Count == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "Por favor, seleccione al menos una semana para generar el reporte.",
                        "Advertencia",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Usar la nueva clase con EPPlus
                Reporte_Semanal_EPPlus reporte = new Reporte_Semanal_EPPlus(connectionString);
                reporte.IncluirCumplimientoMensual = true;

                reporte.ServidorSMTP = servidor_smtp;
                reporte.PuertoSMTP = PuertoSMTP;
                reporte.RemitenteEmail = RemitenteEMail;
                reporte.PasswordEmail = PasswordEmail;
                reporte.DestinatariosEmail = DestinatariosEmail;
                reporte.SSLCheck = SSLCheck;
                string anio_select = CB_Anio_Cumplimiento.Text;
                bool resultado = reporte.GenerarYEnviarReporte(semanasSeleccionadas, anio_select);

                if (resultado)
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "Reporte generado y enviado correctamente por correo electrónico.",
                        "Éxito",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this,
                        "No se pudo completar el envío del reporte.",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    $"Error al generar el reporte: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btn_new_report_costo_Click(object sender, EventArgs e)
        {
            // Validar que la fecha inicial no sea mayor que la final
            if (DTP_Costo_1.Value.Date > DTP_Costo_2.Value.Date)
            {
                MetroFramework.MetroMessageBox.Show(this, "La fecha inicial no puede ser mayor que la fecha final.",
                                "Validación de Fechas",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }

            reporte_costo();
            btn_export_excel_costo.Enabled = true;
            btn_clean_costo.Enabled = true;
            txt_filtro_report_costo.Enabled = true;
            btn_filtro_costo.Enabled = true;
        }
        private void reporte_costo()
        {
            string var1 = cb_area_costo.Text;
            string querySimple = string.Empty;

            // Obtener las fechas de los controles MetroDateTime
            DateTime fechaInicio = DTP_Costo_1.Value.Date;
            DateTime fechaFin = DTP_Costo_2.Value.Date; // Incluye todo el día final

            rgv_reporte_costo.DataSource = null;
            rgv_reporte_costo.Rows.Clear();
            rgv_reporte_costo.Columns.Clear();
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);

            // Crear parámetros para las fechas
            NpgsqlParameter[] parameters = null;

            if (var1 == "Despegue")
            {
                querySimple = @"
SELECT
    ""ID_Ficha"",
    ""Fecha"",
    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
    ""Turno"",
    ""Lote"",
    ""OP"",
    ""Costo""
FROM public.""Ficha""

WHERE ""Area"" = 'Despegue'
AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY ""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Evaporado")
            {
                querySimple = @"
SELECT
    ""ID_Ficha"",
    ""Fecha"",
    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
    ""Turno"",
    ""OP"",
    ""Costo""
FROM public.""Ficha""

WHERE ""Area"" = 'Evaporado'
AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY ""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Grind" || var1 == "Inspeccion" || var1 == "Empacado" || var1 == "Revolturas")
            {
                querySimple = @"
SELECT
    ""ID_Ficha"",
    ""Fecha"",
    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
    ""Turno"",
    ""OP"",
    ""Costo""

FROM public.""Ficha""

WHERE ""Area"" = @Area
AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY ""ID_Ficha"", ""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@Area", NpgsqlTypes.NpgsqlDbType.Varchar) { Value = var1 ?? (object)DBNull.Value },
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Máquinas")
            {
                querySimple = @"
SELECT
    ""ID_Ficha"",
    ""Fecha"",
    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
    ""Turno"",
    ""OP"",
    ""Costo""

FROM public.""Ficha""

WHERE ""Area"" = 'Máquinas'
AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY ""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }
            else if (var1 == "Polvos")
            {
                querySimple = @"
SELECT
    ""ID_Ficha"",
    ""Fecha"",
    EXTRACT(WEEK FROM ""Fecha"") AS ""No. Semana"",
    ""Turno"",
    ""OP"",
    ""Costo""

FROM public.""Ficha""

WHERE ""Area"" = 'Polvos'
AND ""Fecha"" BETWEEN @FechaInicio AND @FechaFin
ORDER BY ""OP"" ASC;";

                parameters = new NpgsqlParameter[]
                {
            new NpgsqlParameter("@FechaInicio", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaInicio },
            new NpgsqlParameter("@FechaFin", NpgsqlTypes.NpgsqlDbType.Date) { Value = fechaFin }
                };
            }

            // Cargar los datos de la tabla Usuarios en el DataGridView
            dbHelper.LoadDataIntoDataGridViewTelerik(querySimple, rgv_reporte_costo, parameters);

            // Verificar si hay datos antes de configurar columnas
            if (rgv_reporte_costo.Columns.Count > 0)
            {
                // Configurar el DataGridView
                rgv_reporte_costo.Columns[0].IsVisible = false; // Ocultar la columna ID
                rgv_reporte_costo.Columns["Fecha"].FormatString = "{0:dd/MM/yyyy}";
                rgv_reporte_costo.Columns["Fecha"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_costo.Columns["No. Semana"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_costo.Columns["Turno"].TextAlignment = ContentAlignment.MiddleCenter;
                rgv_reporte_costo.Columns["OP"].TextAlignment = ContentAlignment.MiddleCenter;
                if (var1 == "Despegue")
                {
                    rgv_reporte_costo.Columns["Lote"].TextAlignment = ContentAlignment.MiddleCenter;
                }
            }
            rgv_reporte_costo.CurrentRow = null;
            // Limpiar el filtro al cargar nuevos datos
            txt_filtro_report_costo.Clear();
        }
        private void cb_area_costo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_area_costo.SelectedIndex != -1)
            {
                btn_new_report_costo.Enabled = true;
                DTP_Costo_1.Enabled = true;
                DTP_Costo_2.Enabled = true;
            }
        }

        private async void btn_export_excel_costo_Click(object sender, EventArgs e)
        {

            LoadingScreen.ShowLoading();
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Guardar archivo de Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    await Task.Run(() =>
                    {
                        ExportarRadGridViewFiltradoAExcel_ClosedXML(
                            rgv_reporte_costo,
                            filePath
                        );
                    });

                    MetroFramework.MetroMessageBox.Show(
                        this,
                        "La exportación fue completada con éxito.",
                        "Exportación a Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(
                    this,
                    "Error durante la exportación: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                LoadingScreen.HideLoading();
            }
        }

        private void btn_clean_costo_Click(object sender, EventArgs e)
        {
            // Limpiar el DataGridView
            rgv_reporte_costo.DataSource = null;
            rgv_reporte_costo.Rows.Clear();
            rgv_reporte_costo.Columns.Clear();
            // Limpiar el filtro al cargar nuevos datos
            txt_filtro_report_costo.Clear();
            // Deshabilitar botones hasta que se seleccione un reporte y área
            btn_export_excel_costo.Enabled = false;
            btn_clean_costo.Enabled = false;
            txt_filtro_report_costo.Enabled = false;
            btn_filtro_costo.Enabled = false;
            btn_new_report_costo.Enabled = false;
            DTP_Costo_1.Enabled = false;
            DTP_Costo_2.Enabled = false;
            cb_area_costo.SelectedIndex = -1;
            //limpiar edicion de costo
            cancel_costo();
        }

        private void btn_filtro_costo_Click(object sender, EventArgs e)
        {
            txt_filtro_report_costo.Clear();
        }

        private void rgv_reporte_costo_CustomFiltering(object sender, GridViewCustomFilteringEventArgs e)
        {
            string textoFiltro = txt_filtro_report_costo.Text.Trim();

            // Si no hay texto de filtro, mostrar todas las filas
            if (string.IsNullOrEmpty(textoFiltro))
            {
                e.Visible = true;
                ResetearEstiloCeldas(e);
                return;
            }

            // Iniciar actualización para mejor rendimiento
            rgv_reporte_costo.BeginUpdate();

            // Por defecto ocultar la fila
            e.Visible = false;

            // Buscar en todas las celdas de la fila
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                GridViewCellInfo celda = e.Row.Cells[i];

                // Verificar si la celda tiene valor
                if (celda.Value != null)
                {
                    string textoCelda = celda.Value.ToString();

                    // Buscar coincidencia (case-insensitive)
                    if (textoCelda.IndexOf(textoFiltro, 0, StringComparison.InvariantCultureIgnoreCase) >= 0)
                    {
                        e.Visible = true; // Mostrar la fila si hay coincidencia

                        // Resaltar la celda que coincide
                        celda.Style.CustomizeFill = true;
                        celda.Style.DrawFill = true;
                        celda.Style.BackColor = Color.FromArgb(201, 252, 254); // Color azul claro
                    }
                    else
                    {
                        // Resetear estilo si no coincide
                        celda.Style.Reset();
                    }
                }
                else
                {
                    // Resetear estilo si la celda es nula
                    celda.Style.Reset();
                }
            }

            rgv_reporte_costo.EndUpdate(false);
        }

        private void txt_filtro_report_costo_TextChanged(object sender, EventArgs e)
        {
            // Refrescar el grid para aplicar el filtro
            if (rgv_reporte_costo != null)
            {
                rgv_reporte_costo.MasterTemplate.Refresh();
            }
        }

        private void rgv_reporte_costo_CellClick(object sender, GridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string var1 = cb_area_costo.Text;

                txt_costo.Enabled = true;

                id_global_costo = rgv_reporte_costo.Rows[e.RowIndex].Cells[0].Value.ToString();
                if (var1 == "Despegue") 
                {
                    txt_costo.Text = rgv_reporte_costo.Rows[e.RowIndex].Cells[6].Value.ToString();
                }else
                {
                    txt_costo.Text = rgv_reporte_costo.Rows[e.RowIndex].Cells[5].Value.ToString();
                }
                // habilitar controles
                btn_save_costo.Enabled = true;
                btn_cancel_costo.Enabled = true;
            }
        }

        private void btn_save_costo_Click(object sender, EventArgs e)
        {
            DatabaseHelper dbHelper = new DatabaseHelper(connectionString);
            string queryInsertUpdate = string.Empty;
            int result;
            if (!string.IsNullOrEmpty(id_global_costo))
            {
                // actualizar
                // Convertir el ID a entero ANTES de crear el parámetro
                int idCosto = System.Convert.ToInt32(id_global_costo);

                queryInsertUpdate = "UPDATE public.\"Ficha\" SET \"Costo\" = @Costo WHERE \"ID_Ficha\" = @ID_Ficha;";

                NpgsqlParameter[] parametersInsertUpdate = new NpgsqlParameter[]
                {
                        new NpgsqlParameter("@Costo", NpgsqlTypes.NpgsqlDbType.Numeric)
                        {
                            Value = System.Convert.ToDecimal(txt_costo.Text)
                        },
                        new NpgsqlParameter("@ID_Ficha", NpgsqlTypes.NpgsqlDbType.Integer)
                        {
                            Value = idCosto  // variable convertida a int
                        }
                };

                result = dbHelper.ExecuteNonQuery(queryInsertUpdate, parametersInsertUpdate);
                reporte_costo();
                cancel_costo();
            }
            else 
            {
                MetroFramework.MetroMessageBox.Show(this, "Por favor, complete todos los campos antes de guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void txt_costo_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite el dígito, el punto decimal y el backspace
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // Bloquea el carácter
                return;
            }

            // Evita que se ingresen múltiples puntos decimales
            if (e.KeyChar == '.' && (sender as RadTextBox).Text.IndexOf('.') > -1)
            {
                e.Handled = true;
                return;
            }
        }

        private void btn_cancel_costo_Click(object sender, EventArgs e)
        {
            cancel_costo();
        }

        private void cancel_costo() 
        {
            // deshabilitar controles
            txt_costo.Enabled = false;
            btn_save_costo.Enabled = false;
            btn_cancel_costo.Enabled = false;

            txt_costo.Text = string.Empty;
            id_global_costo = string.Empty;
            rgv_reporte_costo.CurrentRow = null;
        }

        private void CB_Anio_Cumplimiento_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Obtener el año actual como string
            string anioActual = DateTime.Now.Year.ToString();
            if (anioActual == CB_Anio_Cumplimiento.Text)
            {
                CargarSemanasSimple();
            }
            else 
            {
                CargartodasSemanas(); 
            }
        }
    }
}
