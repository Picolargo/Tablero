using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Tablero.Clases
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Text;
    using System.Threading.Tasks;

    namespace Tablero.Clases
    {
        public class EmailSender
        {
            public static async Task<bool> SendFichaEmail(string area, Dictionary<string, object> datosFicha, EmailConfig config)
            {
                try
                {
                    // Configurar protocolo de seguridad para Office 365
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

                    using (MailMessage mail = new MailMessage())
                    {
                        mail.From = new MailAddress(config.FromEmail);

                        // Obtener lista de destinatarios desde la cadena separada por ;
                        var destinatarios = GetEmailList(config.ToEmail);
                        foreach (var email in destinatarios)
                        {
                            if (IsValidEmail(email))
                            {
                                mail.To.Add(email);
                            }
                        }

                        // Verificar que haya al menos un destinatario
                        if (mail.To.Count == 0)
                        {
                            Console.WriteLine("No hay correos destinatarios válidos configurados.");
                            return false;
                        }

                        mail.Subject = $"📋 Registro de Ficha - Área: {area} - Sistema Tablero";
                        mail.IsBodyHtml = true;
                        mail.Body = CreateHtmlBody(area, datosFicha);

                        using (SmtpClient smtp = new SmtpClient(config.SmtpServer, config.SmtpPort))
                        {
                            // Configuración para Office 365
                            smtp.Credentials = new NetworkCredential(config.SmtpUser, config.SmtpPassword);
                            smtp.EnableSsl = config.EnableSSL;
                            smtp.UseDefaultCredentials = false;
                            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                            smtp.Timeout = 30000; // 30 segundos

                            // Headers adicionales para Office 365
                            mail.Headers.Add("X-Mailer", "SistemaTablero");
                            mail.Headers.Add("Priority", "Normal");

                            await smtp.SendMailAsync(mail);
                            Console.WriteLine($"✅ Correo enviado exitosamente a {mail.To.Count} destinatario(s)");
                            return true;
                        }
                    }
                }
                catch (SmtpException smtpEx)
                {
                    string errorMessage = GetFriendlySmtpError(smtpEx);
                    Console.WriteLine($"❌ Error SMTP: {errorMessage}");
                    Console.WriteLine($"Código de estado: {smtpEx.StatusCode}");

                    // Lanza una excepción con mensaje amigable para mostrar al usuario
                    throw new Exception($"Error al enviar correo:\n{errorMessage}", smtpEx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error enviando correo: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine($"Detalle: {ex.InnerException.Message}");
                    }
                    throw new Exception($"Error al enviar correo:\n{ex.Message}", ex);
                }
            }

            private static string GetFriendlySmtpError(SmtpException ex)
            {
                string message = ex.Message;

                if (message.Contains("5.7.57") || message.Contains("Client not authenticated"))
                {
                    return @"Office 365 requiere configuración adicional:

1️⃣ Habilita SMTP AUTH en Office 365:
   • Portal admin → Usuarios → Mail → Manage email apps
   • Marca 'Authenticated SMTP'

2️⃣ Si tienes autenticación multifactor (MFA):
   • Crea una 'App Password' en: https://mysignins.microsoft.com/security-info
   • Usa esa contraseña especial

3️⃣ Configuración correcta:
   • Servidor: smtp.office365.com
   • Puerto: 587
   • SSL: Habilitado
   • Usuario: correo-completo@dominio.com
   • Contraseña: (normal o App Password)";
                }

                if (message.Contains("535") || message.Contains("Authentication unsuccessful"))
                {
                    return @"Error de autenticación:

🔍 Verifica:
1. Usuario: DEBE ser el correo COMPLETO (ej: productosregionales@picolargo.com)
2. Contraseña: Correcta y actualizada
3. Si tienes MFA, usa 'App Password' no tu contraseña normal
4. La cuenta debe estar activa y no bloqueada";
                }

                if (message.Contains("timed out") || message.Contains("timeout"))
                {
                    return @"Timeout - El servidor no responde:

🔍 Verifica:
1. Conexión a internet
2. Servidor SMTP correcto: smtp.office365.com
3. Puerto: 587
4. Firewall/antivirus no bloquea el puerto 587";
                }

                return $"Error SMTP: {message}";
            }

            // Método para parsear string de correos a lista
            private static List<string> GetEmailList(string emailString)
            {
                var emails = new List<string>();

                if (string.IsNullOrWhiteSpace(emailString))
                    return emails;

                // Separar por punto y coma (como se guarda en settings)
                string[] separators = new[] { ";" };

                foreach (var email in emailString.Split(separators, StringSplitOptions.RemoveEmptyEntries))
                {
                    string cleanEmail = email.Trim();
                    if (!string.IsNullOrWhiteSpace(cleanEmail))
                    {
                        emails.Add(cleanEmail);
                    }
                }

                return emails;
            }

            // Método para validar email
            private static bool IsValidEmail(string email)
            {
                try
                {
                    var addr = new MailAddress(email);
                    return addr.Address == email;
                }
                catch
                {
                    return false;
                }
            }

            private static string CreateHtmlBody(string area, Dictionary<string, object> datos)
            {
                string rows = "";
                foreach (var dato in datos)
                {
                    rows += $@"
            <tr style='border-bottom: 1px solid #e0e0e0;'>
                <td style='padding: 10px; font-weight: bold; color: #555;'>{dato.Key}</td>
                <td style='padding: 10px; color: #333;'>{dato.Value}</td>
            </tr>";
                }

                return $@"
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; }}
                .container {{ max-width: 800px; margin: 0 auto; padding: 20px; }}
                .header {{ background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); color: white; padding: 25px; border-radius: 8px 8px 0 0; }}
                .content {{ background-color: #ffffff; padding: 25px; border-radius: 0 0 8px 8px; border: 1px solid #e0e0e0; }}
                .footer {{ margin-top: 30px; text-align: center; color: #777; font-size: 12px; padding-top: 15px; border-top: 1px solid #eee; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
                th {{ background-color: #f8f9fa; padding: 12px; text-align: left; font-weight: 600; color: #495057; border-bottom: 2px solid #dee2e6; }}
                .note {{ margin-top: 25px; padding: 15px; background-color: #e8f5e8; border-left: 4px solid #4CAF50; border-radius: 4px; }}
            </style>
        </head>
        <body>
            <div class='container'>
                <div class='header'>
                    <h2 style='margin: 0; font-size: 24px;'>📋 Sistema de Tablero - Registro de Ficha</h2>
                    <p style='margin: 5px 0 0 0; font-size: 16px; opacity: 0.9;'>Área: {area} | Fecha: {DateTime.Now:dd/MM/yyyy HH:mm}</p>
                </div>
                
                <div class='content'>
                    <p style='font-size: 16px;'>Estimado equipo,</p>
                    
                    <p>Se ha registrado exitosamente una nueva ficha en el sistema Tablero con los siguientes datos:</p>
                    
                    <table>
                        <thead>
                            <tr>
                                <th style='width: 40%;'>Campo</th>
                                <th style='width: 60%;'>Valor</th>
                            </tr>
                        </thead>
                        <tbody>
                            {rows}
                        </tbody>
                    </table>
                    
                    <div class='note'>
                        <p style='margin: 0;'><strong>📌 Nota:</strong> Este es un correo automático generado por el sistema de Tablero. 
                        Por favor, no responda a este mensaje. Para consultas, contacte al departamento correspondiente.</p>
                    </div>
                </div>
                
                <div class='footer'>
                    <p style='margin: 0;'>© {DateTime.Now.Year} Sistema Tablero - Control de Producción</p>
                    <p style='margin: 5px 0 0 0; font-size: 11px; color: #999;'>
                        Este correo fue enviado automáticamente el {DateTime.Now:dd/MM/yyyy HH:mm:ss}
                    </p>
                </div>
            </div>
        </body>
        </html>";
            }

            // Método para diagnóstico (opcional)
            public static string TestConnection(EmailConfig config)
            {
                try
                {
                    var destinatarios = GetEmailList(config.ToEmail);
                    return $"✅ Configuración cargada:\n" +
                           $"• Servidor: {config.SmtpServer}:{config.SmtpPort}\n" +
                           $"• Usuario: {config.SmtpUser}\n" +
                           $"• SSL: {(config.EnableSSL ? "Habilitado" : "Deshabilitado")}\n" +
                           $"• Remitente: {config.FromEmail}\n" +
                           $"• Destinatarios: {destinatarios.Count} correo(s)";
                }
                catch (Exception ex)
                {
                    return $"❌ Error en configuración: {ex.Message}";
                }
            }
        }
    }
}
