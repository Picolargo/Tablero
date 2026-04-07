using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace Tablero
{
    public partial class Automatico : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Timer checkTimer;
        private string logPath = Path.Combine(Application.StartupPath, "log_semanal.txt");
        private string estadoPath = Path.Combine(Application.StartupPath, "estado_semanal.json");
        private SemanalEstado estado;
        string connectionString = string.Empty;

        private ToolStripMenuItem inicioAutomaticoMenuItem; // Para cambiar el texto dinámicamente

        // Variables para el envío de correos
        private string servidor_smtp = string.Empty;
        private string RemitenteEMail = string.Empty;
        private string PasswordEmail = string.Empty;
        private int PuertoSMTP = 0;
        private string DestinatariosEmail = string.Empty;
        private bool SSLCheck = false;
        // Fin de variables para el envío de correos

        public Automatico(string conexionstring)
        {
            InitializeComponent();
            connectionString = conexionstring;
            // Configurar para que NO sea visible como ventana
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Load += Automatico_Load;
        }

        private void Automatico_Load(object sender, EventArgs e)
        {
            CrearTrayIcon();
            CargarEstado();
            ConfigurarTimer();
            ConfigurarInicioAutomatico();

            // Opcional: mostrar un globo indicando que el servicio inició
            trayIcon.ShowBalloonTip(2000, "Servicio Automático",
                "Monitoreo de tarea semanal iniciado (Lunes 10 AM)",
                ToolTipIcon.Info);
        }

        private void CrearTrayIcon()
        {
            trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "Tablero - Tarea Semanal (Lunes 10 AM)",
                Visible = true
            };

            trayMenu = new ContextMenuStrip();

            // Opciones del menú
            trayMenu.Items.Add("Ejecutar tarea ahora", null, (s, ev) => Automatico_Email("manual"));
            trayMenu.Items.Add("Ver log de ejecuciones", null, (s, ev) => VerLog());
            trayMenu.Items.Add("Reiniciar estado semanal", null, (s, ev) => ReiniciarEstado());
            trayMenu.Items.Add("Configurar Email", null, (s, ev) => AbrirConfiguracionEmail());
            trayMenu.Items.Add("-"); // Separador

            // Opción para habilitar/deshabilitar inicio automático
            inicioAutomaticoMenuItem = new ToolStripMenuItem();
            ActualizarTextoInicioAutomatico();
            inicioAutomaticoMenuItem.Click += (s, ev) => ToggleInicioAutomatico();
            trayMenu.Items.Add(inicioAutomaticoMenuItem);

            trayMenu.Items.Add("-"); // Separador
            trayMenu.Items.Add("Salir", null, (s, ev) => Salir());

            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.DoubleClick += (s, ev) => VerLog();
        }
        private void ActualizarTextoInicioAutomatico()
        {
            bool estaHabilitado = VerificarInicioAutomatico();
            if (estaHabilitado)
            {
                inicioAutomaticoMenuItem.Text = "✓ Deshabilitar inicio automático";
                inicioAutomaticoMenuItem.ForeColor = Color.Green;
            }
            else
            {
                inicioAutomaticoMenuItem.Text = "Habilitar inicio automático";
                inicioAutomaticoMenuItem.ForeColor = Color.Black;
            }
        }

        private bool VerificarInicioAutomatico()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false))
                {
                    if (key != null)
                    {
                        string valor = key.GetValue("TableroAutomatico") as string;
                        return !string.IsNullOrEmpty(valor);
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"[ERROR] Verificando inicio automático: {ex.Message}{Environment.NewLine}");
            }
            return false;
        }

        private void ToggleInicioAutomatico()
        {
            try
            {
                bool estaHabilitado = VerificarInicioAutomatico();

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key != null)
                    {
                        if (estaHabilitado)
                        {
                            // Deshabilitar - eliminar la entrada del registro
                            key.DeleteValue("TableroAutomatico", false);
                            string mensaje = "Inicio automático deshabilitado. El programa no se iniciará con Windows.";
                            trayIcon.ShowBalloonTip(2000, "Configuración cambiada", mensaje, ToolTipIcon.Info);
                            File.AppendAllText(logPath, $"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - Inicio automático deshabilitado{Environment.NewLine}");
                        }
                        else
                        {
                            // Habilitar - agregar la entrada del registro
                            string appPath = Application.ExecutablePath;
                            key.SetValue("TableroAutomatico", appPath);
                            string mensaje = "Inicio automático habilitado. El programa se iniciará con Windows.";
                            trayIcon.ShowBalloonTip(2000, "Configuración cambiada", mensaje, ToolTipIcon.Info);
                            File.AppendAllText(logPath, $"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - Inicio automático habilitado{Environment.NewLine}");
                        }

                        // Actualizar el texto del menú
                        ActualizarTextoInicioAutomatico();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cambiar inicio automático: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                File.AppendAllText(logPath, $"[ERROR] Cambiando inicio automático: {ex.Message}{Environment.NewLine}");
            }
        }
        private void ConfigurarTimer()
        {
            checkTimer = new Timer();
            checkTimer.Interval = 60000; // Revisar cada minuto
            checkTimer.Tick += (s, ev) => RevisarYEjecutar();
            checkTimer.Start();

            // Revisión inmediata al iniciar
            RevisarYEjecutar();
        }

        private void RevisarYEjecutar()
        {
            DateTime ahora = DateTime.Now;
            bool esLunes = ahora.DayOfWeek == DayOfWeek.Monday;
            bool esHorario = ahora.Hour == 10 && ahora.Minute <= 5; // Tolerancia 5 minutos
            string semanaActualId = string.Empty;
            // Si no es Lunes, resetear estado para la próxima semana
            if (!esLunes)
            {
                semanaActualId = ObtenerIdSemana(ahora);
                if (estado.EjecutadoEstaSemana && estado.SemanaId != semanaActualId)
                {
                    estado.EjecutadoEstaSemana = false;
                    estado.FechaEjecucion = null;
                    GuardarEstado();
                }
                return;
            }

            // Es Lunes
            semanaActualId = ObtenerIdSemana(ahora);

            // Si ya se ejecutó esta semana, no hacer nada
            if (estado.EjecutadoEstaSemana && estado.SemanaId == semanaActualId)
                return;

            // Si es horario o ya pasaron las 10 AM (pero no se ejecutó aún)
            if (esHorario || ahora.Hour >= 10)
            {
                Automatico_Email("automático");
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

        private void Automatico_Email(string origen)
        {
            try
            {
                Reporte_Semanal_EPPlus reporte = new Reporte_Semanal_EPPlus(connectionString);
                reporte.IncluirCumplimientoMensual = false;

                // Obtener automáticamente la semana anterior
                List<int> semanasSeleccionadas = reporte.ObtenerSemanaAnterior();

                email_varibles();
                reporte.ServidorSMTP = servidor_smtp;
                reporte.PuertoSMTP = PuertoSMTP;
                reporte.RemitenteEmail = RemitenteEMail;
                reporte.PasswordEmail = PasswordEmail;
                reporte.DestinatariosEmail = DestinatariosEmail;
                reporte.SSLCheck = SSLCheck;
                // Generar y enviar el reporte
                bool resultado = reporte.GenerarYEnviarReporte(semanasSeleccionadas);
                

                // Variable para indicar si tu código se ejecutó correctamente
                bool exito = true; // <--- CAMBIA ESTO SEGÚN EL RESULTADO DE TU CÓDIGO

                // *********************************************************************

                if (exito)
                {
                    DateTime ahora = DateTime.Now;
                    string mensaje = $"[OK] Ejecutado el {ahora:yyyy-MM-dd HH:mm:ss} (origen: {origen})";
                    File.AppendAllText(logPath, mensaje + Environment.NewLine);

                    estado.EjecutadoEstaSemana = true;
                    estado.SemanaId = ObtenerIdSemana(ahora);
                    estado.FechaEjecucion = ahora;
                    GuardarEstado();

                    // Mostrar notificación en bandeja
                    trayIcon.ShowBalloonTip(3000, "Tarea ejecutada correctamente",
                        $"Tarea semanal ejecutada a las {ahora:HH:mm:ss}",
                        ToolTipIcon.Info);
                }
                else
                {
                    string mensajeError = $"[ERROR] Falló el {DateTime.Now:yyyy-MM-dd HH:mm:ss} (origen: {origen}) - Se reintentará";
                    File.AppendAllText(logPath, mensajeError + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"[EXCEPCION] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {ex.Message}{Environment.NewLine}");
            }
        }

        private string ObtenerIdSemana(DateTime fecha)
        {
            // Obtener número de semana (ISO 8601)
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            int semana = culture.Calendar.GetWeekOfYear(
                fecha,
                System.Globalization.CalendarWeekRule.FirstFourDayWeek,
                DayOfWeek.Monday);
            return $"{fecha.Year}-{semana:00}";
        }

        private void CargarEstado()
        {
            if (File.Exists(estadoPath))
            {
                try
                {
                    string json = File.ReadAllText(estadoPath);
                    estado = JsonSerializer.Deserialize<SemanalEstado>(json) ?? new SemanalEstado();
                }
                catch
                {
                    estado = new SemanalEstado();
                }
            }
            else
            {
                estado = new SemanalEstado();
            }
        }

        private void GuardarEstado()
        {
            try
            {
                string json = JsonSerializer.Serialize(estado);
                File.WriteAllText(estadoPath, json);
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"[ERROR] No se pudo guardar estado: {ex.Message}{Environment.NewLine}");
            }
        }

        private void ReiniciarEstado()
        {
            var resultado = MessageBox.Show("¿Reiniciar estado semanal? Esto permitirá ejecutar la tarea nuevamente aunque ya se haya ejecutado esta semana.",
                "Reiniciar Estado", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                estado.EjecutadoEstaSemana = false;
                estado.SemanaId = "";
                estado.FechaEjecucion = null;
                GuardarEstado();
                trayIcon.ShowBalloonTip(2000, "Estado reiniciado",
                    "El estado semanal ha sido reiniciado. Se podrá ejecutar nuevamente.",
                    ToolTipIcon.Info);
            }
        }

        private void VerLog()
        {
            if (File.Exists(logPath))
                System.Diagnostics.Process.Start("notepad.exe", logPath);
            else
                MessageBox.Show("Aún no hay registros de ejecución.", "Log",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ConfigurarInicioAutomatico()
        {
            // Este método ya no configura automáticamente, solo verifica el estado actual
            // La configuración se hace manual desde el menú contextual
            bool estaHabilitado = VerificarInicioAutomatico();
            if (estaHabilitado)
            {
                // Asegurar que la entrada exista (por si alguien la borró manualmente)
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                    {
                        if (key != null)
                        {
                            string appPath = Application.ExecutablePath;
                            string existingValue = key.GetValue("TableroAutomatico") as string;
                            if (string.IsNullOrEmpty(existingValue))
                            {
                                key.SetValue("TableroAutomatico", appPath);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    File.AppendAllText(logPath, $"[ERROR] Configurando inicio automático: {ex.Message}{Environment.NewLine}");
                }
            }
        }

        private void Salir()
        {
            trayIcon.Visible = false;
            Application.Exit();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.Hide(); // Ocultar completamente el formulario
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide(); // Minimizar a bandeja en lugar de cerrar
            }
            else
            {
                base.OnFormClosing(e);
            }
        }
    }

    // Clase para persistir el estado
    public class SemanalEstado
    {
        public bool EjecutadoEstaSemana { get; set; }
        public string SemanaId { get; set; } = "";
        public DateTime? FechaEjecucion { get; set; }
    }
}