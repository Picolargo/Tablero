using System;
using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Win32;

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

        public Automatico()
        {
            InitializeComponent();

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
            // Puedes cambiar el ícono por uno propio
            trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "Tablero - Tarea Semanal (Lunes 10 AM)",
                Visible = true
            };

            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Ejecutar tarea ahora", null, (s, ev) => EjecutarMiCodigo("manual"));
            trayMenu.Items.Add("Ver log de ejecuciones", null, (s, ev) => VerLog());
            trayMenu.Items.Add("Reiniciar estado semanal", null, (s, ev) => ReiniciarEstado());
            trayMenu.Items.Add("-");
            trayMenu.Items.Add("Salir", null, (s, ev) => Salir());

            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.DoubleClick += (s, ev) => VerLog();
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
                EjecutarMiCodigo("automatico");
            }
        }

        private void EjecutarMiCodigo(string origen)
        {
            try
            {
                MessageBox.Show("se ejecuto");
                // *********************************************************************
                // AQUÍ COLOCAS TU CÓDIGO PERSONALIZADO
                // *********************************************************************
                // Ejemplo de lo que puedes hacer:

                // 1. Actualizar alguna tabla en PostgreSQL
                // using var conn = new NpgsqlConnection(connectionString);
                // conn.Open();
                // etc...

                // 2. Generar un reporte
                // 3. Enviar correos automáticos
                // 4. Procesar archivos
                // 5. Llamar a una API

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
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key != null)
                    {
                        string appPath = Application.ExecutablePath;
                        key.SetValue("TableroAutomatico", appPath);
                    }
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, $"[ERROR] No se pudo configurar inicio automático: {ex.Message}{Environment.NewLine}");
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