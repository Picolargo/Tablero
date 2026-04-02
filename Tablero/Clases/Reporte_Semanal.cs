using ClosedXML.Excel;
using System;
using System.Collections.Generic;
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
        /// Obtiene los datos de Evaporado para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataEvaporado(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Evaporado'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene los datos de Grind para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataGrind(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Grind'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene los datos de Inspeccion para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataInspeccion(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Inspeccion'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }
        /// <summary>
        /// Obtiene los datos de Polvos para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataPolvos(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Polvos'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene los datos de Empacado para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataEmpacado(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Empacado'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene los datos de Revolturas para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataRevolturas(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Revolturas'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene los datos de Maquinas para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataMaquinas(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_term"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Máquinas'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_term""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }

        /// <summary>
        /// Obtiene los datos de Deshidratado para las semanas seleccionadas
        /// </summary>
        public DataTable GetDataDeshidratado(List<int> semanasSeleccionadas)
        {
            if (semanasSeleccionadas == null || semanasSeleccionadas.Count == 0)
                return new DataTable();

            string semanasStr = string.Join(",", semanasSeleccionadas);

            string query = $@"
WITH fechas_con_semana AS (
    SELECT 
        f.""Fecha"",
        DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days' as fin_semana,
        f.""Kg_meta"",
        f.""Kg_prod_seco"",
        f.""Kg_fuera_espec""
    FROM public.""Ficha"" f
    WHERE f.""Area"" = 'Despegue'
        AND EXTRACT(YEAR FROM f.""Fecha"") = EXTRACT(YEAR FROM CURRENT_DATE)
),
datos_por_semana AS (
    SELECT 
        EXTRACT(WEEK FROM fin_semana) as semana,
        COALESCE(SUM(f.""Kg_meta""), 0) as meta,
        COALESCE(SUM(f.""Kg_prod_seco""), 0) as producido,
        COALESCE(SUM(f.""Kg_fuera_espec""), 0) as fuera_espec
    FROM fechas_con_semana f
    WHERE EXTRACT(WEEK FROM fin_semana) IN ({semanasStr})
    GROUP BY EXTRACT(WEEK FROM fin_semana)
)
SELECT 
    semana,
    meta,
    producido,
    fuera_espec,
    CASE 
        WHEN meta > 0 THEN 
            LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
        ELSE 0
    END as cumplimiento
FROM datos_por_semana
ORDER BY semana ASC";

            return dbHelper.ExecuteSelectQuery(query);
        }
        public string GenerarExcel(List<int> semanasSeleccionadas)
        {
            try
            {
                // Obtener datos de todas las áreas con las semanas seleccionadas
                DataTable datosEvaporado = GetDataEvaporado(semanasSeleccionadas);
                DataTable datosGrind = GetDataGrind(semanasSeleccionadas);
                DataTable datosInspeccion = GetDataInspeccion(semanasSeleccionadas);
                DataTable datosPolvos = GetDataPolvos(semanasSeleccionadas);
                DataTable datosEmpacado = GetDataEmpacado(semanasSeleccionadas);
                DataTable datosRevolturas = GetDataRevolturas(semanasSeleccionadas);
                DataTable datosMaquinas = GetDataMaquinas(semanasSeleccionadas);
                DataTable datosDeshidratado = GetDataDeshidratado(semanasSeleccionadas);

                // Obtener el número de semana actual
                CultureInfo culture = CultureInfo.CurrentCulture;
                Calendar calendar = culture.Calendar;
                CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
                DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;

                //int semanaActual = calendar.GetWeekOfYear(DateTime.Now, weekRule, firstDayOfWeek);

                // Crear nombre de archivo temporal
                string carpetaTemp = Path.Combine(Path.GetTempPath(), "ReportesTablero");
                if (!Directory.Exists(carpetaTemp))
                    Directory.CreateDirectory(carpetaTemp);

                string archivoExcel = Path.Combine(carpetaTemp, $"Reporte_Cumplimiento.xlsx");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Cumplimiento");

                    // ==================== TABLA 1: ÁREA EVAPORADO ====================
                    int filaInicioEvaporado = 1;
                    GenerarTablaArea(worksheet, datosEvaporado, "EVAPORADO", filaInicioEvaporado);

                    // Calcular la fila donde termina la tabla de Evaporado
                    int filaFinEvaporado = filaInicioEvaporado + 3 + datosEvaporado.Rows.Count;
                    int filaInicioGrind = filaFinEvaporado + 2;

                    // ==================== TABLA 2: ÁREA GRIND ====================
                    GenerarTablaArea(worksheet, datosGrind, "GRIND", filaInicioGrind);

                    int filaFinGrind = filaInicioGrind + 3 + datosGrind.Rows.Count;
                    int filaInicioInspeccion = filaFinGrind + 2;

                    // ==================== TABLA 3: ÁREA INSPECCION ====================
                    GenerarTablaArea(worksheet, datosInspeccion, "Inspeccion", filaInicioInspeccion);

                    int filaFinInspeccion = filaInicioInspeccion + 3 + datosInspeccion.Rows.Count;
                    int filaInicioPolvos = filaFinInspeccion + 2;

                    // ==================== TABLA 4: ÁREA POLVOS ====================
                    GenerarTablaArea(worksheet, datosPolvos, "Polvos", filaInicioPolvos);

                    int filaFinPolvos = filaInicioPolvos + 3 + datosPolvos.Rows.Count;
                    int filaInicioEmpacado = filaFinPolvos + 2;

                    // ==================== TABLA 5: ÁREA EMPACADO ====================
                    GenerarTablaArea(worksheet, datosEmpacado, "Empacado", filaInicioEmpacado);

                    int filaFinEmpacado = filaInicioEmpacado + 3 + datosEmpacado.Rows.Count;
                    int filaInicioRevolturas = filaFinEmpacado + 2;

                    // ==================== TABLA 6: ÁREA Revolturas ====================
                    GenerarTablaArea(worksheet, datosRevolturas, "Revolturas", filaInicioRevolturas);

                    int filaFinRevolturas = filaInicioRevolturas + 3 + datosRevolturas.Rows.Count;
                    int filaInicioMaquinas = filaFinRevolturas + 2;

                    // ==================== TABLA 7: ÁREA Maquinas ====================
                    GenerarTablaArea(worksheet, datosMaquinas, "Máquinas", filaInicioMaquinas);

                    int filaFinMaquinas = filaInicioMaquinas + 3 + datosMaquinas.Rows.Count;
                    int filaInicioDeshidratado = filaFinMaquinas + 2;

                    // ==================== TABLA 8: ÁREA Deshidratado ====================
                    GenerarTablaArea(worksheet, datosDeshidratado, "Deshidratado", filaInicioDeshidratado);

                    // ==================== TABLA 9: PORCENTAJE DE CUMPLIMIENTO SEMANAL ====================
                    GenerarTablaCumplimientoSemanal(worksheet,
                        datosEvaporado, datosGrind, datosInspeccion,
                        datosPolvos, datosEmpacado, datosRevolturas,
                        datosMaquinas, datosDeshidratado, semanasSeleccionadas);

                    // Ajustar automáticamente el ancho de las columnas
                    worksheet.Columns().AdjustToContents();

                    // Configurar página para impresión
                    worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    worksheet.PageSetup.PaperSize = XLPaperSize.LetterPaper;
                    worksheet.PageSetup.Margins.Left = 0.5;
                    worksheet.PageSetup.Margins.Right = 0.5;
                    worksheet.PageSetup.Margins.Top = 0.5;
                    worksheet.PageSetup.Margins.Bottom = 0.5;

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
        /// Método auxiliar para generar una tabla de un área específica
        /// </summary>
        private void GenerarTablaArea(IXLWorksheet worksheet, DataTable datos, string nombreArea, int filaInicio)
        {
            int filaActual = filaInicio;

            // 1. Título del área
            worksheet.Cell(filaActual, 1).Value = $"REPORTE SEMANAL - ÁREA {nombreArea}";
            worksheet.Cell(filaActual, 1).Style.Font.SetBold(true);
            worksheet.Cell(filaActual, 1).Style.Font.SetFontSize(16);
            worksheet.Cell(filaActual, 1).Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
            worksheet.Cell(filaActual, 1).Style.Fill.SetBackgroundColor(ColoresReporte.ColorTitulo);
            worksheet.Range(filaActual, 1, filaActual, 5).Merge();
            worksheet.Range(filaActual, 1, filaActual, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range(filaActual, 1, filaActual, 5).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            worksheet.Row(filaActual).Height = 30;
            filaActual++;

            // 2. Subtítulo con fecha de generación
            worksheet.Cell(filaActual, 1).Value = $"Fecha de generación: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
            worksheet.Cell(filaActual, 1).Style.Font.SetFontSize(10);
            worksheet.Cell(filaActual, 1).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
            worksheet.Range(filaActual, 1, filaActual, 5).Merge();
            filaActual++;

            // 3. Encabezados de la tabla
            string[] encabezados = { "Semana", "Meta (kg)", "Producido (kg)", "Fuera de Espec. (kg)", "Cumplimiento (%)" };
            for (int i = 0; i < encabezados.Length; i++)
            {
                var celda = worksheet.Cell(filaActual, i + 1);
                celda.Value = encabezados[i];
                celda.Style.Font.SetBold(true);
                celda.Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
                celda.Style.Fill.SetBackgroundColor(ColoresReporte.ColorEncabezado);
                celda.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                celda.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                celda.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                celda.Style.Border.SetOutsideBorderColor(XLColor.Gray);
            }
            worksheet.Row(filaActual).Height = 25;
            filaActual++;

            // 4. Datos de la tabla
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

                // Color alternativo para filas
                if (i % 2 == 1)
                {
                    worksheet.Range(filaActual, 1, filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorFondoAlternativo);
                }

                // Color condicional para cumplimiento
                if (cumplimiento >= 0.90m)
                {
                    worksheet.Cell(filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoAlto);
                    worksheet.Cell(filaActual, 5).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                }
                else if (cumplimiento >= 0.80m)
                {
                    worksheet.Cell(filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoMedio);
                    worksheet.Cell(filaActual, 5).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                }
                else
                {
                    worksheet.Cell(filaActual, 5).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoBajo);
                    worksheet.Cell(filaActual, 5).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
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
        }
        /// <summary>
        /// Genera la tabla de porcentaje de cumplimiento semanal combinado
        /// </summary>
        private void GenerarTablaCumplimientoSemanal(IXLWorksheet worksheet,
            DataTable datosEvaporado, DataTable datosGrind, DataTable datosInspeccion,
            DataTable datosPolvos, DataTable datosEmpacado, DataTable datosRevolturas,
            DataTable datosMaquinas, DataTable datosDeshidratado, List<int> semanasSeleccionadas)
        {
            // Usar las semanas seleccionadas directamente
            var todasSemanas = semanasSeleccionadas.OrderBy(s => s).ToList();

            // Crear un diccionario para acceder rápidamente a los valores de cumplimiento por área y semana
            var cumplimientoDeshidratado = new Dictionary<int, decimal>();
            foreach (DataRow row in datosDeshidratado.Rows)
            {
                int semana = Convert.ToInt32(row["semana"]);
                decimal cumplimiento = Convert.ToDecimal(row["cumplimiento"]);
                cumplimientoDeshidratado[semana] = cumplimiento;
            }

            // Diccionario para almacenar cumplimientos de todas las áreas (excepto Deshidratado)
            var areasCumplimiento = new Dictionary<int, List<decimal>>();

            void AgregarCumplimiento(DataTable tabla, Dictionary<int, List<decimal>> dict)
            {
                foreach (DataRow row in tabla.Rows)
                {
                    int semana = Convert.ToInt32(row["semana"]);
                    decimal cumplimiento = Convert.ToDecimal(row["cumplimiento"]);

                    if (!dict.ContainsKey(semana))
                        dict[semana] = new List<decimal>();

                    dict[semana].Add(cumplimiento);
                }
            }

            AgregarCumplimiento(datosEvaporado, areasCumplimiento);
            AgregarCumplimiento(datosGrind, areasCumplimiento);
            AgregarCumplimiento(datosInspeccion, areasCumplimiento);
            AgregarCumplimiento(datosPolvos, areasCumplimiento);
            AgregarCumplimiento(datosEmpacado, areasCumplimiento);
            AgregarCumplimiento(datosRevolturas, areasCumplimiento);
            AgregarCumplimiento(datosMaquinas, areasCumplimiento);

            // Configurar la posición inicial de la tabla (columna G = columna 7)
            int columnaInicio = 7;
            int filaActual = 1;

            // 1. Título de la tabla
            worksheet.Cell(filaActual, columnaInicio).Value = "PORCENTAJE DE CUMPLIMIENTO SEMANAL";
            worksheet.Cell(filaActual, columnaInicio).Style.Font.SetBold(true);
            worksheet.Cell(filaActual, columnaInicio).Style.Font.SetFontSize(14);
            worksheet.Cell(filaActual, columnaInicio).Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
            worksheet.Cell(filaActual, columnaInicio).Style.Fill.SetBackgroundColor(ColoresReporte.ColorTitulo);
            worksheet.Cell(filaActual, columnaInicio).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Cell(filaActual, columnaInicio).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            worksheet.Cell(filaActual, columnaInicio).Style.Alignment.SetWrapText(true);
            worksheet.Range(filaActual, columnaInicio, filaActual, columnaInicio + 3).Merge();
            worksheet.Range(filaActual, columnaInicio, filaActual, columnaInicio + 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range(filaActual, columnaInicio, filaActual, columnaInicio + 3).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            worksheet.Row(filaActual).Height = 30;
            filaActual++;

            // 2. Encabezados de la tabla
            string[] encabezadosTabla = { "Semana", "Deshidratado", "Demás Áreas", "Total" };
            for (int i = 0; i < encabezadosTabla.Length; i++)
            {
                var celda = worksheet.Cell(filaActual, columnaInicio + i);
                celda.Value = encabezadosTabla[i];
                celda.Style.Font.SetBold(true);
                celda.Style.Font.SetFontColor(ColoresReporte.ColorFuenteBlanca);
                celda.Style.Fill.SetBackgroundColor(ColoresReporte.ColorEncabezado);
                celda.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                celda.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                celda.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                celda.Style.Border.SetOutsideBorderColor(XLColor.Gray);
            }
            worksheet.Row(filaActual).Height = 25;
            filaActual++;

            // 3. Datos de la tabla por cada semana seleccionada
            foreach (int semana in todasSemanas)
            {
                // Verificar si Deshidratado tiene valor
                bool tieneDeshidratado = cumplimientoDeshidratado.ContainsKey(semana);

                // Calcular Deshidratado (si tiene valor, se multiplica por 50%)
                decimal valorDeshidratado = 0;
                if (tieneDeshidratado)
                {
                    valorDeshidratado = cumplimientoDeshidratado[semana] * 0.5m;
                }

                // Calcular Demás Áreas
                decimal valorDemasAreas = 0;
                if (areasCumplimiento.ContainsKey(semana) && areasCumplimiento[semana].Count > 0)
                {
                    decimal suma = 0;
                    foreach (var cumplimiento in areasCumplimiento[semana])
                    {
                        suma += cumplimiento;
                    }
                    decimal promedioDemasAreas = suma / areasCumplimiento[semana].Count;

                    // Si Deshidratado tiene valor, se multiplica por 50%, si no, se muestra el promedio directo
                    if (tieneDeshidratado)
                    {
                        valorDemasAreas = promedioDemasAreas * 0.5m;
                    }
                    else
                    {
                        valorDemasAreas = promedioDemasAreas;
                    }
                }

                // Calcular Total
                decimal valorTotal = valorDeshidratado + valorDemasAreas;

                // Semana
                worksheet.Cell(filaActual, columnaInicio).Value = semana;
                worksheet.Cell(filaActual, columnaInicio).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(filaActual, columnaInicio).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                // Deshidratado
                if (valorDeshidratado > 0)
                {
                    worksheet.Cell(filaActual, columnaInicio + 1).Value = valorDeshidratado / 100;
                    worksheet.Cell(filaActual, columnaInicio + 1).Style.NumberFormat.SetFormat("0.00%");
                }
                else
                {
                    worksheet.Cell(filaActual, columnaInicio + 1).Value = "";
                }
                worksheet.Cell(filaActual, columnaInicio + 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(filaActual, columnaInicio + 1).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                // Demás Áreas
                if (valorDemasAreas > 0)
                {
                    worksheet.Cell(filaActual, columnaInicio + 2).Value = valorDemasAreas / 100;
                    worksheet.Cell(filaActual, columnaInicio + 2).Style.NumberFormat.SetFormat("0.00%");
                }
                else
                {
                    worksheet.Cell(filaActual, columnaInicio + 2).Value = "";
                }
                worksheet.Cell(filaActual, columnaInicio + 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(filaActual, columnaInicio + 2).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                // Total con formato condicional
                if (valorTotal > 0)
                {
                    decimal cumplimientoPorcentaje = valorTotal / 100;

                    decimal valorMostrar = cumplimientoPorcentaje > 1.0m ? 1.0m : cumplimientoPorcentaje;

                    worksheet.Cell(filaActual, columnaInicio + 3).Value = valorMostrar;
                    worksheet.Cell(filaActual, columnaInicio + 3).Style.NumberFormat.SetFormat("0.00%");

                    if (cumplimientoPorcentaje >= 0.90m)
                    {
                        worksheet.Cell(filaActual, columnaInicio + 3).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoAlto);
                        worksheet.Cell(filaActual, columnaInicio + 3).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                    }
                    else if (cumplimientoPorcentaje >= 0.80m)
                    {
                        worksheet.Cell(filaActual, columnaInicio + 3).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoMedio);
                        worksheet.Cell(filaActual, columnaInicio + 3).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                    }
                    else
                    {
                        worksheet.Cell(filaActual, columnaInicio + 3).Style.Fill.SetBackgroundColor(ColoresReporte.ColorCumplimientoBajo);
                        worksheet.Cell(filaActual, columnaInicio + 3).Style.Font.SetFontColor(ColoresReporte.ColorFuenteOscura);
                    }

                    worksheet.Cell(filaActual, columnaInicio + 3).Style.Font.SetBold(true);
                }
                else
                {
                    worksheet.Cell(filaActual, columnaInicio + 3).Value = "";
                }
                worksheet.Cell(filaActual, columnaInicio + 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(filaActual, columnaInicio + 3).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                // Color alternativo para filas
                if (Array.IndexOf(todasSemanas.ToArray(), semana) % 2 == 1)
                {
                    worksheet.Range(filaActual, columnaInicio, filaActual, columnaInicio + 3).Style.Fill.SetBackgroundColor(ColoresReporte.ColorFondoAlternativo);
                }

                filaActual++;
            }

            // Ajustar el ancho de las columnas de la tabla
            for (int i = 0; i < 4; i++)
            {
                worksheet.Column(columnaInicio + i).AdjustToContents();
                if (worksheet.Column(columnaInicio + i).Width < 12)
                {
                    worksheet.Column(columnaInicio + i).Width = 12;
                }
            }
        }
        /// <summary>
        /// Envía el reporte por correo electrónico
        /// </summary>
        public bool EnviarReportePorCorreo(string archivoExcel, List<int> semanasSeleccionadas)
        {
            try
            {
                if (string.IsNullOrEmpty(archivoExcel) || !File.Exists(archivoExcel))
                    throw new Exception("El archivo Excel no existe o no se pudo generar.");

                string semanasStr = string.Join(",", semanasSeleccionadas);
                string cuerpoHtml;
                // Cuerpo HTML del correo
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
                                <h2>Reporte de Cumplimiento Semanal</h2>
                            </div>
                            <div class='content'>
                                <p>Estimado usuario,</p>
                                <p>Se adjunta el reporte general semanal de porcentajes de cumplimiento correspondiente a las semanas {semanasStr}.</p>
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

                    

                    // Usarlo en el asunto
                    correo.Subject = $"Reporte de Cumplimiento generado la fecha {DateTime.Now:dd/MM/yyyy}";
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
        /// Método principal que genera el reporte con las semanas seleccionadas y lo envía por correo
        /// </summary>
        public bool GenerarYEnviarReporte(List<int> semanasSeleccionadas)
        {
            string archivoExcel = null;
            try
            {
                // Generar Excel con las semanas seleccionadas
                archivoExcel = GenerarExcel(semanasSeleccionadas);

                // Enviar por correo
                bool enviado = EnviarReportePorCorreo(archivoExcel, semanasSeleccionadas);

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
    }
}