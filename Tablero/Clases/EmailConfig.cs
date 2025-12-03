using System;
using System.Collections.Generic;

namespace Tablero.Clases
{
    public class EmailConfig
    {
        public string SmtpServer { get; set; }
        public int SmtpPort { get; set; }
        public string SmtpUser { get; set; }
        public string SmtpPassword { get; set; }
        public bool EnableSSL { get; set; }
        public string FromEmail { get; set; }
        public string ToEmail { get; set; } // Cadena con correos separados por ;

        public static EmailConfig LoadFromSettings()
        {
            return new EmailConfig
            {
                SmtpServer = Properties.Settings.Default.SmtpServer ?? "smtp.office365.com",
                SmtpPort = Properties.Settings.Default.SmtpPort != 0 ?
                          Properties.Settings.Default.SmtpPort : 587,
                SmtpUser = Properties.Settings.Default.SmtpUser ?? "productosregionales@picolargo.com",
                SmtpPassword = Properties.Settings.Default.SmtpPassword ?? "Picolargo1234",
                EnableSSL = Properties.Settings.Default.EnableSSL,
                FromEmail = Properties.Settings.Default.FromEmail ?? "productosregionales@picolargo.com",
                ToEmail = Properties.Settings.Default.ToEmail ?? ""
            };
        }

        // Método para obtener lista de destinatarios
        public List<string> GetRecipientList()
        {
            var recipients = new List<string>();

            if (string.IsNullOrWhiteSpace(this.ToEmail))
                return recipients;

            // Separar por punto y coma
            string[] emails = this.ToEmail.Split(';');

            foreach (string email in emails)
            {
                string cleanEmail = email.Trim();
                if (!string.IsNullOrWhiteSpace(cleanEmail))
                {
                    recipients.Add(cleanEmail);
                }
            }

            return recipients;
        }
    }
}
