using ClosedXML.Excel;
using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace Tablero
{
    public class Reporte_Semanal
    {
        private DatabaseHelper dbHelper;
        private string connectionString;

        // Propiedades de configuración de correo
        public string ServidorSMTP { get; set; }
        public int PuertoSMTP { get; set; }
        public string RemitenteEmail { get; set; }
        public string PasswordEmail { get; set; }
        public string DestinatariosEmail { get; set; }
        public bool SSLCheck { get; set; }

        // Configuración del reporte
        public string NombreReporte { get; set; } = "Reporte Semanal";
        public int NumeroSemanas { get; set; } = 4;

        // Áreas disponibles para reporte
        public enum AreaReporte
        {
            Evaporado,
            Deshidratado,
            Grind,
            Inspeccion,
            Polvos,
            Empacado,
            Revolturas,
            Maquinas,
            Todos
        }

        public Reporte_Semanal(string connectionString)
        {
            this.connectionString = connectionString;
            this.dbHelper = new DatabaseHelper(connectionString);
            InicializarConfiguracionPorDefecto();
        }

        private void InicializarConfiguracionPorDefecto()
        {
            // Configuración de colores profesional
            ColoresReporte = new ReporteColores
            {
                ColorTitulo = XLColor.FromArgb(52, 73, 94),      // Azul oscuro profesional
                ColorEncabezado = XLColor.FromArgb(41, 128, 185), // Azul pastel
                ColorFondoAlternativo = XLColor.FromArgb(235, 245, 251), // Azul muy claro
                ColorFuenteBlanca = XLColor.White,
                ColorFuenteOscura = XLColor.FromArgb(52, 73, 94),
                ColorCumplimientoAlto = XLColor.FromArgb(39, 174, 96),   // Verde
                ColorCumplimientoMedio = XLColor.FromArgb(241, 196, 15),  // Amarillo
                ColorCumplimientoBajo = XLColor.FromArgb(231, 76, 60)     // Rojo
            };
        }

        // Estructura para colores del reporte
        public class ReporteColores
        {
            public XLColor ColorTitulo { get; set; }
            public XLColor ColorEncabezado { get; set; }
            public XLColor ColorFondoAlternativo { get; set; }
            public XLColor ColorFuenteBlanca { get; set; }
            public XLColor ColorFuenteOscura { get; set; }
            public XLColor ColorCumplimientoAlto { get; set; }
            public XLColor ColorCumplimientoMedio { get; set; }
            public XLColor ColorCumplimientoBajo { get; set; }
        }

        public ReporteColores ColoresReporte { get; set; }

        /// <summary>
        /// Obtiene los datos de Evaporado para las últimas N semanas
        /// </summary>
        public DataTable GetData()
        {
            string query = $@"
                WITH semanas AS (
                    SELECT DISTINCT 
                        EXTRACT(WEEK FROM ""Fecha"") as semana,
                        DATE_TRUNC('week', ""Fecha"") as inicio_semana
                    FROM public.""Ficha""
                    WHERE ""Area"" = 'Evaporado'
                    ORDER BY inicio_semana DESC
                    LIMIT {NumeroSemanas}
                ),
                datos_semana AS (
                    SELECT 
                        EXTRACT(WEEK FROM f.""Fecha"") as semana,
                        COALESCE(SUM(e.""Meta_kg_hr""), 0) as meta,
                        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
                        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
                    FROM public.""Ficha"" f
                    INNER JOIN public.""Evaporado"" e ON f.""OP"" = e.""OP""
                    WHERE f.""Area"" = 'Evaporado'
                        AND EXTRACT(WEEK FROM f.""Fecha"") IN (SELECT semana FROM semanas)
                    GROUP BY EXTRACT(WEEK FROM f.""Fecha"")
                )
                SELECT 
                    s.semana,
                    COALESCE(ds.meta, 0) as meta,
                    COALESCE(ds.producido, 0) as producido,
                    COALESCE(ds.fuera_espec, 0) as fuera_espec,
                    CASE 
                        WHEN COALESCE(ds.meta, 0) > 0 THEN 
                            LEAST(ROUND(((COALESCE(ds.producido, 0) - COALESCE(ds.fuera_espec, 0)) / COALESCE(ds.meta, 0)) * 100, 2),100)
                        ELSE 0
                    END as cumplimiento
                FROM semanas s
                LEFT JOIN datos_semana ds ON s.semana = ds.semana
                ORDER BY s.semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Método genérico para obtener datos de cualquier área
        /// </summary>
        public DataTable GetAreaData(AreaReporte area)
        {
            string tablaArea = ObtenerNombreTabla(area);
            if (string.IsNullOrEmpty(tablaArea))
                return null;

            // Obtener el campo de meta según el área
            string campoMeta = ObtenerCampoMeta(area);

            string query = $@"
                WITH semanas AS (
                    SELECT DISTINCT 
                        EXTRACT(WEEK FROM ""Fecha"") as semana,
                        DATE_TRUNC('week', ""Fecha"") as inicio_semana
                    FROM public.""Ficha""
                    WHERE ""Area"" = '{area.ToString()}'
                    ORDER BY inicio_semana DESC
                    LIMIT {NumeroSemanas}
                ),
                datos_semana AS (
                    SELECT 
                        EXTRACT(WEEK FROM f.""Fecha"") as semana,
                        COALESCE(SUM(a.{campoMeta}), 0) as meta,
                        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
                        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
                    FROM public.""Ficha"" f
                    INNER JOIN public.""{tablaArea}"" a ON f.""OP"" = a.""OP""
                    WHERE f.""Area"" = '{area.ToString()}'
                        AND EXTRACT(WEEK FROM f.""Fecha"") IN (SELECT semana FROM semanas)
                    GROUP BY EXTRACT(WEEK FROM f.""Fecha"")
                )
                SELECT 
                    s.semana,
                    COALESCE(ds.meta, 0) as meta,
                    COALESCE(ds.producido, 0) as producido,
                    COALESCE(ds.fuera_espec, 0) as fuera_espec,
                    CASE 
                        WHEN COALESCE(ds.meta, 0) > 0 THEN 
                            LEAST(ROUND(((COALESCE(ds.producido, 0) - COALESCE(ds.fuera_espec, 0)) / COALESCE(ds.meta, 0)) * 100, 2),100)
                        ELSE 0
                    END as cumplimiento
                FROM semanas s
                LEFT JOIN datos_semana ds ON s.semana = ds.semana
                ORDER BY s.semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene el nombre de la tabla según el área
        /// </summary>
        private string ObtenerNombreTabla(AreaReporte area)
        {
            switch (area)
            {
                case AreaReporte.Evaporado:
                    return "Evaporado";
                case AreaReporte.Deshidratado:
                    return "Deshidratado";
                case AreaReporte.Grind:
                    return "Grind";
                case AreaReporte.Inspeccion:
                    return "Inspeccion";
                case AreaReporte.Polvos:
                    return "Polvos";
                case AreaReporte.Empacado:
                    return "Empacado";
                case AreaReporte.Revolturas:
                    return "Revolturas";
                case AreaReporte.Maquinas:
                    return "Maquinas";
                default:
                    return null;
            }
        }

        /// <summary>
        /// Obtiene el campo de meta según el área
        /// </summary>
        private string ObtenerCampoMeta(AreaReporte area)
        {
            switch (area)
            {
                case AreaReporte.Evaporado:
                    return "\"Meta_kg_hr\"";
                case AreaReporte.Deshidratado:
                    return "\"Kg_seco_hr\"";
                case AreaReporte.Grind:
                    return "\"Meta_Kg_hr\"";
                case AreaReporte.Inspeccion:
                    return "\"Meta_kg_hr_line\"";
                case AreaReporte.Polvos:
                    return "\"Meta_kg_hr_idon\"";
                case AreaReporte.Empacado:
                    return "\"Meta_kg_hr_line\"";
                case AreaReporte.Revolturas:
                    return "\"Meta_Kg_Hr\"";
                case AreaReporte.Maquinas:
                    return "\"Meta_Kg_Hr\"";
                default:
                    return "\"Meta_kg_hr\"";
            }
        }

        /// <summary>
        /// Genera el archivo Excel con formato profesional para un área específica
        /// </summary>
        //public string GenerarExcel(AreaReporte area)
        //{
        //    DataTable datos = GetAreaData(area);
        //    return GenerarExcelBase(datos);
        //}

        /// <summary>
        /// Genera el archivo Excel con formato profesional para Evaporado (mantener compatibilidad)
        /// </summary>
        public string GenerarExcel()
        {
            DataTable datos = GetData();
            return GenerarExcelBase(datos);
        }

        /// <summary>
        /// Método base para generar el Excel con formato profesional
        /// </summary>
        private string GenerarExcelBase(DataTable datos)
        {
            
            try
            {
                
                // Obtener el número de semana actual
                CultureInfo culture = CultureInfo.CurrentCulture;
                Calendar calendar = culture.Calendar;
                CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
                DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;

                int semanaActual = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstDayOfWeek);

                // Crear nombre de archivo temporal
                string carpetaTemp = Path.Combine(Path.GetTempPath(), "ReportesTablero");
                if (!Directory.Exists(carpetaTemp))
                    Directory.CreateDirectory(carpetaTemp);

                string archivoExcel = Path.Combine(carpetaTemp, $"Reporte_semanal_No_{semanaActual}.xlsx");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Cumplimiento");
                    //////////////////////////////////////////////////////////////////REPORTES SEMANALES - ÁREA EVAPORADO//////////////////////////////////////////////////////////////////////////////////////////////////////
                    // 1. Título del reporte
                    worksheet.Cell("A1").Value = $"REPORTE SEMANAL - ÁREA EVAPORADO";
                    worksheet.Cell("A1").Style.Font.SetBold(true);
                    worksheet.Cell("A1").Style.Font.SetFontSize(16);
                    worksheet.Cell("A1").Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
                    worksheet.Cell("A1").Style.Fill.SetBackgroundColor(ColoresReporte.ColorTitulo);
                    worksheet.Range("A1:E1").Merge();
                    worksheet.Range("A1:E1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    worksheet.Range("A1:E1").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                    worksheet.Row(1).Height = 30;

                    // 2. Subtítulo con fecha de generación
                    worksheet.Cell("A2").Value = $"Fecha de generación: {DateTime.Now:dd/MM/yyyy HH:mm:ss} | Últimas {NumeroSemanas} semanas";
                    worksheet.Cell("A2").Style.Font.SetFontSize(10);
                    worksheet.Cell("A2").Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                    worksheet.Range("A2:E2").Merge();

                    // 3. Encabezados de la tabla
                    string[] encabezados = { "Semana", "Meta (kg)", "Producido (kg)", "Fuera de Espec. (kg)", "Cumplimiento (%)" };
                    for (int i = 0; i < encabezados.Length; i++)
                    {
                        var celda = worksheet.Cell(3, i + 1);
                        celda.Value = encabezados[i];
                        celda.Style.Font.SetBold(true);
                        celda.Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
                        celda.Style.Fill.SetBackgroundColor(ColoresReporte.ColorEncabezado);
                        celda.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        celda.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                        celda.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                        celda.Style.Border.SetOutsideBorderColor(XLColor.Gray);
                    }
                    worksheet.Row(3).Height = 25;

                    // 4. Datos de la tabla
                    int filaActual = 4;
                    for (int i = 0; i < datos.Rows.Count; i++)
                    {
                        var fila = datos.Rows[i];

                        // Semana
                        worksheet.Cell(filaActual, 1).Value = Convert.ToInt32(fila["semana"]);
                        worksheet.Cell(filaActual, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                        // Meta
                        worksheet.Cell(filaActual, 2).Value = Convert.ToDecimal(fila["meta"]);
                        worksheet.Cell(filaActual, 2).Style.NumberFormat.SetFormat("#,##0.00");
                        worksheet.Cell(filaActual, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                        // Producido
                        worksheet.Cell(filaActual, 3).Value = Convert.ToDecimal(fila["producido"]);
                        worksheet.Cell(filaActual, 3).Style.NumberFormat.SetFormat("#,##0.00");
                        worksheet.Cell(filaActual, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                        // Fuera de Especificación
                        worksheet.Cell(filaActual, 4).Value = Convert.ToDecimal(fila["fuera_espec"]);
                        worksheet.Cell(filaActual, 4).Style.NumberFormat.SetFormat("#,##0.00");
                        worksheet.Cell(filaActual, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);

                        // Cumplimiento
                        decimal cumplimiento = Convert.ToDecimal(fila["cumplimiento"]) / 100;
                        worksheet.Cell(filaActual, 5).Value = cumplimiento;
                        worksheet.Cell(filaActual, 5).Style.NumberFormat.SetFormat("0.00%");
                        worksheet.Cell(filaActual, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        if (i % 2 == 1)
                        {
                            worksheet.Range(filaActual, 1, filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorFondoAlternativo);
                        }
                        // Color condicional para cumplimiento
                        if (cumplimiento >= 0.90m) 
                        {
                            worksheet.Cell(filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoAlto);
                            worksheet.Cell(filaActual, 5).Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
                        }
                        else if (cumplimiento >= 0.80m)
                        {
                            worksheet.Cell(filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoMedio);
                            worksheet.Cell(filaActual, 5).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                        }
                        else
                        {
                            worksheet.Cell(filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoBajo);
                            worksheet.Cell(filaActual, 5).Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
                        }
                            
                        worksheet.Cell(filaActual, 5).Style.Font.SetBold(true);

                        // Bordes para las celdas de datos
                        for (int col = 1; col <= 5; col++)
                        {
                            worksheet.Cell(filaActual, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                            worksheet.Cell(filaActual, col).Style.Border.SetOutsideBorderColor(XLColor.LightGray);
                        }

                        filaActual++;
                    }

                    // 6. Ajustar automáticamente el ancho de las columnas
                    worksheet.Columns().AdjustToContents();

                    // 7. Configurar página para impresión (versión corregida)
                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    worksheet.PageSetup.PaperSize = XLPaperSize.LetterPaper;
                    worksheet.PageSetup.Margins.Left = 0.5;
                    worksheet.PageSetup.Margins.Right = 0.5;
                    worksheet.PageSetup.Margins.Top = 0.5;
                    worksheet.PageSetup.Margins.Bottom = 0.5;
                    //////////////////////////////////////////////////////////////////////////REPORTES SEMANALES - ÁREA GRIND//////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///

                    

                    // Guardar archivo
                    workbook.SaveAs(archivoExcel);
                }

                return archivoExcel;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar el archivo Excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Envía el reporte por correo electrónico
        /// </summary>
        public bool EnviarReportePorCorreo(string archivoExcel, string area = null, string cuerpoHtml = null)
        {
            try
            {
                if (string.IsNullOrEmpty(archivoExcel) || !File.Exists(archivoExcel))
                    throw new Exception("El archivo Excel no existe o no se pudo generar.");

                //string nombreArea = area ?? "Evaporado";

                // Cuerpo HTML del correo
                if (string.IsNullOrEmpty(cuerpoHtml))
                {
                    cuerpoHtml = $@"
                        <html>
                        <head>
                            <style>
                                body {{ font-family: Arial, sans-serif; color: #333; }}
                                .header {{ background-color: #2c3e50; color: white; padding: 20px; text-align: center; }}
                                .content {{ padding: 20px; }}
                                .footer {{ background-color: #f8f9fa; padding: 10px; text-align: center; font-size: 12px; color: #666; margin-top: 20px; }}
                                table {{ border-collapse: collapse; width: 100%; margin: 15px 0; }}
                                th {{ background-color: #2980b9; color: white; padding: 10px; text-align: center; }}
                                td {{ padding: 8px; border: 1px solid #ddd; }}
                                .highlight {{ background-color: #e8f4fd; }}
                            </style>
                        </head>
                        <body>
                            <div class='header'>
                                <h2>Reporte Automático de Cumplimiento Semanal</h2>
                            </div>
                            <div class='content'>
                                <p>Estimado usuario,</p>
                                <p>Se adjunta el reporte general semanal de porcentajes de cumplimiento correspondiente a las últimas {NumeroSemanas} semanas.</p>
                                <p><strong>Fecha de generación:</strong> {DateTime.Now:dd/MM/yyyy HH:mm:ss}</p>
                                <p>El reporte contiene la siguiente información por semana:</p>
                                <ul>
                                    <li>Meta de producción (kg)</li>
                                    <li>Producción real (kg)</li>
                                    <li>Producto fuera de especificación (kg)</li>
                                    <li>Porcentaje de cumplimiento</li>
                                </ul>
                                <p>Por favor, revise el archivo adjunto para más detalles.</p>
                                <p>Este es un mensaje generado automáticamente por el Sistema Tablero.</p>
                            </div>
                            <div class='footer'>
                                <p>Sistema Tablero - Reporte Automático</p>
                                <p>Este correo fue enviado automáticamente, por favor no responder.</p>
                            </div>
                        </body>
                        </html>";
                }

                // Configurar correo
                using (MailMessage correo = new MailMessage())
                {
                    correo.From = new MailAddress(RemitenteEmail);

                    string[] destinatarios = DestinatariosEmail.Split(',');
                    foreach (var mail in destinatarios)
                    {
                        if (!string.IsNullOrWhiteSpace(mail))
                            correo.To.Add(mail.Trim());
                    }
                    // Obtener el número de semana actual
                    CultureInfo culture = CultureInfo.CurrentCulture;
                    Calendar calendar = culture.Calendar;
                    CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
                    DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;

                    int semanaActual = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstDayOfWeek);

                    // Usarlo en el asunto
                    correo.Subject = $"Reporte Semana No. {semanaActual} - {DateTime.Now:dd/MM/yyyy}";
                    correo.Body = cuerpoHtml;
                    correo.IsBodyHtml = true;

                    // Adjuntar archivo Excel
                    using (Attachment attachment = new Attachment(archivoExcel))
                    {
                        correo.Attachments.Add(attachment);

                        // Configurar SMTP
                        using (SmtpClient smtp = new SmtpClient(ServidorSMTP, PuertoSMTP))
                        {
                            smtp.Credentials = new NetworkCredential(RemitenteEmail, PasswordEmail);
                            smtp.EnableSsl = SSLCheck;

                            // Enviar correo
                            smtp.Send(correo);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al enviar el correo: {ex.Message}");
            }
        }

        /// <summary>
        /// Método principal que genera el reporte de Evaporado y lo envía por correo
        /// </summary>
        public bool GenerarYEnviarReporte()
        {
            string archivoExcel = null;
            try
            {
                // Generar Excel
                archivoExcel = GenerarExcel();

                // Enviar por correo
                bool enviado = EnviarReportePorCorreo(archivoExcel, "Evaporado");

                return enviado;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar y enviar el reporte: {ex.Message}");
            }
            finally
            {
                // Eliminar archivo temporal después de enviar
                if (archivoExcel != null && File.Exists(archivoExcel))
                {
                    try
                    {
                        File.Delete(archivoExcel);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Método genérico para generar y enviar reporte de cualquier área
        /// </summary>
        //public bool GenerarYEnviarReporte(AreaReporte area)
        //{
        //    string archivoExcel = null;
        //    try
        //    {
        //        // Generar Excel
        //        archivoExcel = GenerarExcel(area);

        //        // Enviar por correo
        //        bool enviado = EnviarReportePorCorreo(archivoExcel, area.ToString());

        //        return enviado;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception($"Error al generar y enviar el reporte de {area}: {ex.Message}");
        //    }
        //    finally
        //    {
        //        // Eliminar archivo temporal después de enviar
        //        if (archivoExcel != null && File.Exists(archivoExcel))
        //        {
        //            try
        //            {
        //                File.Delete(archivoExcel);
        //            }
        //            catch { }
        //        }
        //    }
        //}
    }
}