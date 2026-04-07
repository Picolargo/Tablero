using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace Tablero
{
    public class Reporte_Semanal_EPPlus
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
        public bool IncluirCumplimientoMensual { get; set; } = true;
        // Colores profesionales
        private readonly Color ColorTitulo = Color.FromArgb(52, 73, 94);
        private readonly Color ColorEncabezado = Color.FromArgb(41, 128, 185);
        private readonly Color ColorFondoAlternativo = Color.FromArgb(235, 245, 251);
        private readonly Color ColorFuenteBlanca = Color.White;
        private readonly Color ColorFuenteOscura = Color.FromArgb(52, 73, 94);
        private readonly Color ColorCumplimientoAlto = Color.FromArgb(39, 174, 96);
        private readonly Color ColorCumplimientoMedio = Color.FromArgb(241, 196, 15);
        private readonly Color ColorCumplimientoBajo = Color.FromArgb(231, 76, 60);

        public Reporte_Semanal_EPPlus(string connectionString)
        {
            this.connectionString = connectionString;
            this.dbHelper = new DatabaseHelper(connectionString);
        }
        // Método helper para obtener la semana anterior automáticamente
        public List<int> ObtenerSemanaAnterior()
        {
            var listaSemanas = new List<int>();

            // Obtener la fecha actual y calcular la semana anterior
            DateTime fechaActual = DateTime.Now;
            // Ir al lunes de la semana actual y restar 7 días para obtener el lunes de la semana anterior
            int diff = (fechaActual.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)fechaActual.DayOfWeek) - 1;
            DateTime lunesSemanaActual = fechaActual.AddDays(-diff).Date;
            DateTime lunesSemanaAnterior = lunesSemanaActual.AddDays(-7);

            // Obtener el número de semana
            CultureInfo ci = CultureInfo.CurrentCulture;
            Calendar cal = ci.Calendar;
            int semanaNum = cal.GetWeekOfYear(lunesSemanaAnterior, ci.DateTimeFormat.CalendarWeekRule, ci.DateTimeFormat.FirstDayOfWeek);

            listaSemanas.Add(semanaNum);
            return listaSemanas;
        }

        #region Métodos de obtención de datos (iguales a tu clase original)

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

        #endregion
        /// <summary>
        /// Genera la tabla PROMEDIO DEMÁS ÁREAS
        /// </summary>
        private int GenerarTablaPromedioDemasAreasEPPlus(ExcelWorksheet worksheet,
            DataTable datosEvaporado, DataTable datosGrind, DataTable datosInspeccion,
            DataTable datosPolvos, DataTable datosEmpacado, DataTable datosRevolturas,
            DataTable datosMaquinas, DataTable datosDeshidratado, List<int> semanasSeleccionadas, int filaInicio)
        {
            var todasSemanas = semanasSeleccionadas.OrderBy(s => s).ToList();

            // Obtener el valor de "Demás Áreas" de la tabla "PORCENTAJE DE CUMPLIMIENTO SEMANAL"
            var valorDemasAreasPorSemana = new Dictionary<int, decimal>();

            // Primero calculamos los valores de "Demás Áreas" igual que en GenerarTablaCumplimientoSemanalEPPlus
            var cumplimientoDeshidratado = new Dictionary<int, decimal>();
            foreach (DataRow row in datosDeshidratado.Rows)
            {
                int semana = Convert.ToInt32(row["semana"]);
                decimal cumplimiento = Convert.ToDecimal(row["cumplimiento"]);
                cumplimientoDeshidratado[semana] = cumplimiento;
            }

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

            // Calcular el valor de "Demás Áreas" para cada semana (igual que en la tabla)
            foreach (int semana in todasSemanas)
            {
                bool tieneDeshidratado = cumplimientoDeshidratado.ContainsKey(semana);

                decimal valorDemasAreas = 0;
                if (areasCumplimiento.ContainsKey(semana) && areasCumplimiento[semana].Count > 0)
                {
                    decimal suma = 0;
                    foreach (var cumplimiento in areasCumplimiento[semana])
                    {
                        suma += cumplimiento;
                    }
                    decimal promedioDemasAreas = suma / areasCumplimiento[semana].Count;

                    if (tieneDeshidratado)
                    {
                        // Si hay Deshidratado, el valor se multiplicó por 50%
                        valorDemasAreas = promedioDemasAreas * 0.5m;
                    }
                    else
                    {
                        valorDemasAreas = promedioDemasAreas;
                    }
                }
                valorDemasAreasPorSemana[semana] = valorDemasAreas;
            }

            int filaActual = filaInicio+3;
            int colInicio = 12; // Columna L

            // Título de la tabla
            worksheet.Cells[filaActual, colInicio].Value = "PROMEDIO DEMÁS ÁREAS";
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Merge = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.Font.Bold = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.Font.Size = 12;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.Font.Color.SetColor(ColorFuenteBlanca);
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.Fill.BackgroundColor.SetColor(ColorTitulo);
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 1].Style.WrapText = true;
            worksheet.Row(filaActual).Height = 30;
            filaActual++;

            // Encabezados
            worksheet.Cells[filaActual, colInicio].Value = "Semana";
            worksheet.Cells[filaActual, colInicio].Style.Font.Bold = true;
            worksheet.Cells[filaActual, colInicio].Style.Font.Color.SetColor(ColorFuenteBlanca);
            worksheet.Cells[filaActual, colInicio].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[filaActual, colInicio].Style.Fill.BackgroundColor.SetColor(ColorEncabezado);
            worksheet.Cells[filaActual, colInicio].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[filaActual, colInicio].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

            worksheet.Cells[filaActual, colInicio + 1].Value = "Promedio";
            worksheet.Cells[filaActual, colInicio + 1].Style.Font.Bold = true;
            worksheet.Cells[filaActual, colInicio + 1].Style.Font.Color.SetColor(ColorFuenteBlanca);
            worksheet.Cells[filaActual, colInicio + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[filaActual, colInicio + 1].Style.Fill.BackgroundColor.SetColor(ColorEncabezado);
            worksheet.Cells[filaActual, colInicio + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[filaActual, colInicio + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            worksheet.Row(filaActual).Height = 25;
            filaActual++;

            // Datos
            foreach (int semana in todasSemanas)
            {
                bool tieneDeshidratado = cumplimientoDeshidratado.ContainsKey(semana);

                // Obtener el valor de "Demás Áreas" de la tabla
                decimal valorDemasAreas = valorDemasAreasPorSemana.ContainsKey(semana) ? valorDemasAreasPorSemana[semana] : 0;

                // Calcular el promedio original
                decimal promedioOriginal = 0;

                if (tieneDeshidratado)
                {
                    // Si había Deshidratado, el valor en "Demás Áreas" ya está multiplicado por 50%
                    // Para obtener el promedio original, dividimos entre 0.5 (multiplicamos por 2)
                    promedioOriginal = valorDemasAreas / 0.5m;
                }
                else
                {
                    // Si no había Deshidratado, el valor es directo
                    promedioOriginal = valorDemasAreas;
                }

                // Limitar a 100% máximo
                if (promedioOriginal > 100)
                    promedioOriginal = 100;

                // Semana
                worksheet.Cells[filaActual, colInicio].Value = semana;
                worksheet.Cells[filaActual, colInicio].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[filaActual, colInicio].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                // Promedio (convertir a porcentaje dividiendo entre 100)
                worksheet.Cells[filaActual, colInicio + 1].Value = promedioOriginal / 100;
                worksheet.Cells[filaActual, colInicio + 1].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[filaActual, colInicio + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[filaActual, colInicio + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                // Color condicional para el promedio
                decimal cumplimientoPorcentaje = promedioOriginal / 100;
                if (cumplimientoPorcentaje >= 0.90m)
                {
                    worksheet.Cells[filaActual, colInicio + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[filaActual, colInicio + 1].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoAlto);
                    worksheet.Cells[filaActual, colInicio + 1].Style.Font.Color.SetColor(ColorFuenteOscura);
                }
                else if (cumplimientoPorcentaje >= 0.80m)
                {
                    worksheet.Cells[filaActual, colInicio + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[filaActual, colInicio + 1].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoMedio);
                    worksheet.Cells[filaActual, colInicio + 1].Style.Font.Color.SetColor(ColorFuenteOscura);
                }
                else
                {
                    worksheet.Cells[filaActual, colInicio + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[filaActual, colInicio + 1].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoBajo);
                    worksheet.Cells[filaActual, colInicio + 1].Style.Font.Color.SetColor(ColorFuenteOscura);
                }

                worksheet.Cells[filaActual, colInicio + 1].Style.Font.Bold = true;

                // Color alternativo para filas
                if (Array.IndexOf(todasSemanas.ToArray(), semana) % 2 == 1)
                {
                    for (int c = 0; c < 2; c++)
                    {
                        worksheet.Cells[filaActual, colInicio + c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[filaActual, colInicio + c].Style.Fill.BackgroundColor.SetColor(ColorFondoAlternativo);
                    }
                }

                filaActual++;
            }

            return filaActual;
        }
        #region Métodos para obtener datos de la gráfica

        private System.Data.DataTable ObtenerDatosParaGrafica(
            DataTable datosEvaporado, DataTable datosGrind, DataTable datosInspeccion,
            DataTable datosPolvos, DataTable datosEmpacado, DataTable datosRevolturas,
            DataTable datosMaquinas, DataTable datosDeshidratado, List<int> semanasSeleccionadas)
        {
            var cumplimientoDeshidratado = new Dictionary<int, decimal>();
            foreach (DataRow row in datosDeshidratado.Rows)
            {
                int semana = Convert.ToInt32(row["semana"]);
                decimal cumplimiento = Convert.ToDecimal(row["cumplimiento"]);
                cumplimientoDeshidratado[semana] = cumplimiento;
            }

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

            var dataTable = new System.Data.DataTable();
            dataTable.Columns.Add("Semana", typeof(int));
            dataTable.Columns.Add("Total", typeof(decimal));

            foreach (int semana in semanasSeleccionadas.OrderBy(s => s))
            {
                bool tieneDeshidratado = cumplimientoDeshidratado.ContainsKey(semana);

                decimal valorDeshidratado = 0;
                if (tieneDeshidratado)
                {
                    valorDeshidratado = cumplimientoDeshidratado[semana] * 0.5m;
                }

                decimal valorDemasAreas = 0;
                if (areasCumplimiento.ContainsKey(semana) && areasCumplimiento[semana].Count > 0)
                {
                    decimal suma = 0;
                    foreach (var cumplimiento in areasCumplimiento[semana])
                    {
                        suma += cumplimiento;
                    }
                    decimal promedioDemasAreas = suma / areasCumplimiento[semana].Count;

                    if (tieneDeshidratado)
                    {
                        valorDemasAreas = promedioDemasAreas * 0.5m;
                    }
                    else
                    {
                        valorDemasAreas = promedioDemasAreas;
                    }
                }

                decimal valorTotal = valorDeshidratado + valorDemasAreas;
                valorTotal = valorTotal / 100;

                dataTable.Rows.Add(semana, valorTotal);
            }

            return dataTable;
        }

        #endregion
        /// <summary>
        /// Calcula el promedio de la columna Total de la tabla PORCENTAJE DE CUMPLIMIENTO SEMANAL
        /// </summary>
        private decimal CalcularPromedioMensual(
            DataTable datosEvaporado, DataTable datosGrind, DataTable datosInspeccion,
            DataTable datosPolvos, DataTable datosEmpacado, DataTable datosRevolturas,
            DataTable datosMaquinas, DataTable datosDeshidratado, List<int> semanasSeleccionadas)
        {
            var todasSemanas = semanasSeleccionadas.OrderBy(s => s).ToList();

            var cumplimientoDeshidratado = new Dictionary<int, decimal>();
            foreach (DataRow row in datosDeshidratado.Rows)
            {
                int semana = Convert.ToInt32(row["semana"]);
                decimal cumplimiento = Convert.ToDecimal(row["cumplimiento"]);
                cumplimientoDeshidratado[semana] = cumplimiento;
            }

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

            decimal sumaTotales = 0;
            int contadorTotales = 0;

            foreach (int semana in todasSemanas)
            {
                bool tieneDeshidratado = cumplimientoDeshidratado.ContainsKey(semana);

                decimal valorDeshidratado = 0;
                if (tieneDeshidratado)
                    valorDeshidratado = cumplimientoDeshidratado[semana] * 0.5m;

                decimal valorDemasAreas = 0;
                if (areasCumplimiento.ContainsKey(semana) && areasCumplimiento[semana].Count > 0)
                {
                    decimal suma = 0;
                    foreach (var cumplimiento in areasCumplimiento[semana])
                        suma += cumplimiento;
                    decimal promedioDemasAreas = suma / areasCumplimiento[semana].Count;

                    if (tieneDeshidratado)
                        valorDemasAreas = promedioDemasAreas * 0.5m;
                    else
                        valorDemasAreas = promedioDemasAreas;
                }

                decimal valorTotal = valorDeshidratado + valorDemasAreas;
                if (valorTotal > 0)
                {
                    sumaTotales += valorTotal;
                    contadorTotales++;
                }
            }

            return contadorTotales > 0 ? sumaTotales / contadorTotales : 0;
        }
        #region Métodos para generar el Excel con EPPlus

        public string GenerarExcel(List<int> semanasSeleccionadas)
        {
            try
            {
                // Configurar licencia de EPPlus
                ExcelPackage.License.SetNonCommercialPersonal("ReporteSemanal");

                // Obtener datos
                DataTable datosEvaporado = GetDataEvaporado(semanasSeleccionadas);
                DataTable datosGrind = GetDataGrind(semanasSeleccionadas);
                DataTable datosInspeccion = GetDataInspeccion(semanasSeleccionadas);
                DataTable datosPolvos = GetDataPolvos(semanasSeleccionadas);
                DataTable datosEmpacado = GetDataEmpacado(semanasSeleccionadas);
                DataTable datosRevolturas = GetDataRevolturas(semanasSeleccionadas);
                DataTable datosMaquinas = GetDataMaquinas(semanasSeleccionadas);
                DataTable datosDeshidratado = GetDataDeshidratado(semanasSeleccionadas);

                System.Data.DataTable datosGrafica = ObtenerDatosParaGrafica(
                    datosEvaporado, datosGrind, datosInspeccion,
                    datosPolvos, datosEmpacado, datosRevolturas,
                    datosMaquinas, datosDeshidratado, semanasSeleccionadas);

                string carpetaTemp = Path.Combine(Path.GetTempPath(), "ReportesTablero");
                if (!Directory.Exists(carpetaTemp))
                    Directory.CreateDirectory(carpetaTemp);

                string archivoExcel = Path.Combine(carpetaTemp, $"Reporte_Cumplimiento_EPPlus.xlsx");

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Cumplimiento");

                    // Generar tablas
                    int filaActual = 1;
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosEvaporado, "EVAPORADO", filaActual);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosGrind, "GRIND", filaActual + 2);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosInspeccion, "Inspeccion", filaActual + 2);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosPolvos, "Polvos", filaActual + 2);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosEmpacado, "Empacado", filaActual + 2);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosRevolturas, "Revolturas", filaActual + 2);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosMaquinas, "Máquinas", filaActual + 2);
                    filaActual = GenerarTablaAreaEPPlus(worksheet, datosDeshidratado, "Deshidratado", filaActual + 2);

                    filaActual = GenerarTablaCumplimientoSemanalEPPlus(worksheet,
                        datosEvaporado, datosGrind, datosInspeccion,
                        datosPolvos, datosEmpacado, datosRevolturas,
                        datosMaquinas, datosDeshidratado, semanasSeleccionadas, 26);

                    // ==================== TABLA PROMEDIO DEMÁS ÁREAS ====================
                    // La tabla comenzará en la celda L26 (columna 12, fila 26)
                    filaActual = GenerarTablaPromedioDemasAreasEPPlus(worksheet,
                        datosEvaporado, datosGrind, datosInspeccion,
                        datosPolvos, datosEmpacado, datosRevolturas,
                        datosMaquinas, datosDeshidratado, semanasSeleccionadas, 26);

                    if (IncluirCumplimientoMensual) 
                    {
                        // Calcular promedio mensual
                        decimal promedioMensual = CalcularPromedioMensual(
                        datosEvaporado, datosGrind, datosInspeccion,
                        datosPolvos, datosEmpacado, datosRevolturas,
                        datosMaquinas, datosDeshidratado, semanasSeleccionadas);
                        decimal promedioMensualMostrar = promedioMensual / 100;

                        Color verdePastel = Color.FromArgb(198, 224, 180);

                        // Título "CUMPLIMIENTO MENSUAL" desde G6 hasta M6
                        worksheet.Cells[6, 7].Value = "CUMPLIMIENTO MENSUAL";
                        worksheet.Cells[6, 7, 6, 13].Merge = true;
                        worksheet.Cells[6, 7, 6, 13].Style.Font.Bold = true;
                        worksheet.Cells[6, 7, 6, 13].Style.Font.Size = 12;
                        worksheet.Cells[6, 7, 6, 13].Style.Font.Color.SetColor(Color.Black);
                        worksheet.Cells[6, 7, 6, 13].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[6, 7, 6, 13].Style.Fill.BackgroundColor.SetColor(verdePastel);
                        worksheet.Cells[6, 7, 6, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[6, 7, 6, 13].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        worksheet.Row(6).Height = 25;

                        // Valor del promedio en N6
                        worksheet.Cells[6, 14].Value = promedioMensualMostrar;
                        worksheet.Cells[6, 14].Style.Numberformat.Format = "0.00%";
                        worksheet.Cells[6, 14].Style.Font.Bold = true;
                        worksheet.Cells[6, 14].Style.Font.Size = 12;
                        worksheet.Cells[6, 14].Style.Font.Color.SetColor(Color.Black);
                        worksheet.Cells[6, 14].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[6, 14].Style.Fill.BackgroundColor.SetColor(verdePastel);
                        worksheet.Cells[6, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[6, 14].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    // Título de la gráfica
                    worksheet.Cells[3, 7].Value = "GRÁFICA DE CUMPLIMIENTO SEMANAL";
                    worksheet.Cells[3, 7, 3, 14].Merge = true;
                    worksheet.Cells[3, 7, 3, 14].Style.Font.Bold = true;
                    worksheet.Cells[3, 7, 3, 14].Style.Font.Size = 12;
                    worksheet.Cells[3, 7, 3, 14].Style.Font.Color.SetColor(ColorFuenteBlanca);
                    worksheet.Cells[3, 7, 3, 14].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[3, 7, 3, 14].Style.Fill.BackgroundColor.SetColor(ColorTitulo);
                    worksheet.Cells[3, 7, 3, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Row(8).Height = 25;

                    // Subtítulo
                    worksheet.Cells[4, 7].Value = $"Reporte generado el {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                    worksheet.Cells[4, 7, 4, 14].Merge = true;
                    worksheet.Cells[4, 7, 4, 14].Style.Font.Size = 8;
                    worksheet.Cells[4, 7, 4, 14].Style.Font.Color.SetColor(ColorFuenteOscura);
                    worksheet.Cells[4, 7, 4, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    // Generar gráfica
                    GenerarGraficaEPPlus(worksheet, datosGrafica, 7, 6);

                    // Autoajustar columnas
                    worksheet.Cells[1, 1, filaActual + 20, 20].AutoFitColumns();

                    package.SaveAs(new FileInfo(archivoExcel));
                }

                return archivoExcel;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar el archivo Excel: {ex.Message}");
            }
        }

        private int GenerarTablaAreaEPPlus(ExcelWorksheet worksheet, DataTable datos, string nombreArea, int filaInicio)
        {
            if (datos.Rows.Count == 0)
                return filaInicio-2;

            int filaActual = filaInicio;
            int colInicio = 1;

            // Título
            worksheet.Cells[filaActual, colInicio].Value = $"REPORTE SEMANAL - ÁREA {nombreArea}";
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Merge = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Font.Bold = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Font.Size = 16;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Font.Color.SetColor(ColorFuenteBlanca);
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Fill.BackgroundColor.SetColor(ColorTitulo);
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Row(filaActual).Height = 30;
            filaActual++;

            // Subtítulo
            worksheet.Cells[filaActual, colInicio].Value = $"Fecha de generación: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Merge = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Font.Size = 10;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 4].Style.Font.Color.SetColor(ColorFuenteOscura);
            filaActual++;

            // Encabezados
            string[] encabezados = { "Semana", "Meta (kg)", "Producido (kg)", "Fuera de Espec. (kg)", "Cumplimiento (%)" };
            for (int i = 0; i < encabezados.Length; i++)
            {
                worksheet.Cells[filaActual, colInicio + i].Value = encabezados[i];
                worksheet.Cells[filaActual, colInicio + i].Style.Font.Bold = true;
                worksheet.Cells[filaActual, colInicio + i].Style.Font.Color.SetColor(ColorFuenteBlanca);
                worksheet.Cells[filaActual, colInicio + i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[filaActual, colInicio + i].Style.Fill.BackgroundColor.SetColor(ColorEncabezado);
                worksheet.Cells[filaActual, colInicio + i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[filaActual, colInicio + i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }
            worksheet.Row(filaActual).Height = 25;
            filaActual++;

            // Datos
            for (int i = 0; i < datos.Rows.Count; i++)
            {
                var fila = datos.Rows[i];

                worksheet.Cells[filaActual, colInicio].Value = Convert.ToInt32(fila["semana"]);
                worksheet.Cells[filaActual, colInicio + 1].Value = Convert.ToDecimal(fila["meta"]);
                worksheet.Cells[filaActual, colInicio + 2].Value = Convert.ToDecimal(fila["producido"]);
                worksheet.Cells[filaActual, colInicio + 3].Value = Convert.ToDecimal(fila["fuera_espec"]);
                decimal cumplimiento = Convert.ToDecimal(fila["cumplimiento"]) / 100;
                worksheet.Cells[filaActual, colInicio + 4].Value = cumplimiento;

                // Formatos
                worksheet.Cells[filaActual, colInicio + 1].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[filaActual, colInicio + 2].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[filaActual, colInicio + 3].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[filaActual, colInicio + 4].Style.Numberformat.Format = "0.00%";
                worksheet.Cells[filaActual, colInicio + 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                // Color alternativo para filas
                if (i % 2 == 1)
                {
                    for (int c = 0; c < 5; c++)
                    {
                        worksheet.Cells[filaActual, colInicio + c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[filaActual, colInicio + c].Style.Fill.BackgroundColor.SetColor(ColorFondoAlternativo);
                    }
                }

                // Color condicional para cumplimiento
                if (cumplimiento >= 0.90m)
                {
                    worksheet.Cells[filaActual, colInicio + 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[filaActual, colInicio + 4].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoAlto);
                    worksheet.Cells[filaActual, colInicio + 4].Style.Font.Color.SetColor(ColorFuenteBlanca);
                }
                else if (cumplimiento >= 0.80m)
                {
                    worksheet.Cells[filaActual, colInicio + 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[filaActual, colInicio + 4].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoMedio);
                    worksheet.Cells[filaActual, colInicio + 4].Style.Font.Color.SetColor(ColorFuenteOscura);
                }
                else
                {
                    worksheet.Cells[filaActual, colInicio + 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[filaActual, colInicio + 4].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoBajo);
                    worksheet.Cells[filaActual, colInicio + 4].Style.Font.Color.SetColor(ColorFuenteBlanca);
                }

                worksheet.Cells[filaActual, colInicio + 4].Style.Font.Bold = true;

                // Bordes
                for (int c = 0; c < 5; c++)
                {
                    worksheet.Cells[filaActual, colInicio + c].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }

                filaActual++;
            }

            return filaActual;
        }

        private int GenerarTablaCumplimientoSemanalEPPlus(ExcelWorksheet worksheet,
            DataTable datosEvaporado, DataTable datosGrind, DataTable datosInspeccion,
            DataTable datosPolvos, DataTable datosEmpacado, DataTable datosRevolturas,
            DataTable datosMaquinas, DataTable datosDeshidratado, List<int> semanasSeleccionadas, int filaInicio)
        {
            var todasSemanas = semanasSeleccionadas.OrderBy(s => s).ToList();

            var cumplimientoDeshidratado = new Dictionary<int, decimal>();
            foreach (DataRow row in datosDeshidratado.Rows)
            {
                int semana = Convert.ToInt32(row["semana"]);
                decimal cumplimiento = Convert.ToDecimal(row["cumplimiento"]);
                cumplimientoDeshidratado[semana] = cumplimiento;
            }

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

            int filaActual = filaInicio+3;
            int colInicio = 7;

            // Título
            worksheet.Cells[filaActual, colInicio].Value = "PORCENTAJE DE CUMPLIMIENTO SEMANAL";
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Merge = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.Font.Bold = true;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.Font.Size = 14;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.Font.Color.SetColor(ColorFuenteBlanca);
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.Fill.BackgroundColor.SetColor(ColorTitulo);
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells[filaActual, colInicio, filaActual, colInicio + 3].Style.WrapText = true;
            worksheet.Row(filaActual).Height = 30;
            filaActual++;

            // Encabezados
            string[] encabezados = { "Semana", "Deshidratado", "Demás Áreas", "Total" };
            for (int i = 0; i < encabezados.Length; i++)
            {
                worksheet.Cells[filaActual, colInicio + i].Value = encabezados[i];
                worksheet.Cells[filaActual, colInicio + i].Style.Font.Bold = true;
                worksheet.Cells[filaActual, colInicio + i].Style.Font.Color.SetColor(ColorFuenteBlanca);
                worksheet.Cells[filaActual, colInicio + i].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[filaActual, colInicio + i].Style.Fill.BackgroundColor.SetColor(ColorEncabezado);
                worksheet.Cells[filaActual, colInicio + i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[filaActual, colInicio + i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }
            worksheet.Row(filaActual).Height = 25;
            filaActual++;

            // Datos
            foreach (int semana in todasSemanas)
            {
                bool tieneDeshidratado = cumplimientoDeshidratado.ContainsKey(semana);

                decimal valorDeshidratado = 0;
                if (tieneDeshidratado)
                {
                    valorDeshidratado = cumplimientoDeshidratado[semana] * 0.5m;
                }

                decimal valorDemasAreas = 0;
                if (areasCumplimiento.ContainsKey(semana) && areasCumplimiento[semana].Count > 0)
                {
                    decimal suma = 0;
                    foreach (var cumplimiento in areasCumplimiento[semana])
                    {
                        suma += cumplimiento;
                    }
                    decimal promedioDemasAreas = suma / areasCumplimiento[semana].Count;

                    if (tieneDeshidratado)
                    {
                        valorDemasAreas = promedioDemasAreas * 0.5m;
                    }
                    else
                    {
                        valorDemasAreas = promedioDemasAreas;
                    }
                }

                decimal valorTotal = valorDeshidratado + valorDemasAreas;

                worksheet.Cells[filaActual, colInicio].Value = semana;
                if (valorDeshidratado > 0)
                    worksheet.Cells[filaActual, colInicio + 1].Value = valorDeshidratado / 100;
                if (valorDemasAreas > 0)
                    worksheet.Cells[filaActual, colInicio + 2].Value = valorDemasAreas / 100;
                if (valorTotal > 0)
                    worksheet.Cells[filaActual, colInicio + 3].Value = valorTotal / 100;

                // Formatos
                for (int i = 1; i <= 3; i++)
                {
                    worksheet.Cells[filaActual, colInicio + i].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[filaActual, colInicio + i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }

                // Color condicional para Total
                if (valorTotal > 0)
                {
                    decimal cumplimientoPorcentaje = valorTotal / 100;
                    worksheet.Cells[filaActual, colInicio + 3].Style.Font.Bold = true;

                    if (cumplimientoPorcentaje >= 0.90m)
                    {
                        worksheet.Cells[filaActual, colInicio + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[filaActual, colInicio + 3].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoAlto);
                        worksheet.Cells[filaActual, colInicio + 3].Style.Font.Color.SetColor(ColorFuenteOscura);
                    }
                    else if (cumplimientoPorcentaje >= 0.80m)
                    {
                        worksheet.Cells[filaActual, colInicio + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[filaActual, colInicio + 3].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoMedio);
                        worksheet.Cells[filaActual, colInicio + 3].Style.Font.Color.SetColor(ColorFuenteOscura);
                    }
                    else
                    {
                        worksheet.Cells[filaActual, colInicio + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[filaActual, colInicio + 3].Style.Fill.BackgroundColor.SetColor(ColorCumplimientoBajo);
                        worksheet.Cells[filaActual, colInicio + 3].Style.Font.Color.SetColor(ColorFuenteOscura);
                    }
                }

                // Color alternativo para filas
                if (Array.IndexOf(todasSemanas.ToArray(), semana) % 2 == 1)
                {
                    for (int c = 0; c < 4; c++)
                    {
                        worksheet.Cells[filaActual, colInicio + c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[filaActual, colInicio + c].Style.Fill.BackgroundColor.SetColor(ColorFondoAlternativo);
                    }
                }

                // Bordes
                for (int c = 0; c < 4; c++)
                {
                    worksheet.Cells[filaActual, colInicio + c].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }

                filaActual++;
            }

            return filaActual;
        }

        //private void GenerarGraficaEPPlus(ExcelWorksheet worksheet, DataTable datosGrafica, int filaInicio, int columnaInicio)
        //{
        //    if (datosGrafica.Rows.Count == 0) return;

        //    var tempSheet = worksheet.Workbook.Worksheets.Add("TempData");
        //    tempSheet.Cells[1, 1].Value = "Semana";
        //    tempSheet.Cells[1, 2].Value = "Cumplimiento";
        //    tempSheet.Cells[1, 1].Style.Font.Bold = true;
        //    tempSheet.Cells[1, 2].Style.Font.Bold = true;

        //    for (int i = 0; i < datosGrafica.Rows.Count; i++)
        //    {
        //        tempSheet.Cells[i + 2, 1].Value = Convert.ToInt32(datosGrafica.Rows[i]["Semana"]);
        //        tempSheet.Cells[i + 2, 2].Value = Convert.ToDecimal(datosGrafica.Rows[i]["Total"]);
        //        tempSheet.Cells[i + 2, 2].Style.Numberformat.Format = "0%";
        //    }

        //    var chart = worksheet.Drawings.AddChart("CumplimientoChart", eChartType.ColumnClustered);
        //    chart.SetPosition(filaInicio, 0, columnaInicio, 0);
        //    chart.SetSize(600, 400);

        //    var serie = chart.Series.Add(tempSheet.Cells[2, 2, datosGrafica.Rows.Count + 1, 2],
        //                                 tempSheet.Cells[2, 1, datosGrafica.Rows.Count + 1, 1]) as ExcelBarChartSerie;

        //    if (serie != null)
        //    {
        //        serie.Header = "Cumplimiento Semanal";

        //        serie.DataLabel.ShowValue = true;
        //    }

        //    chart.Title.Text = "Cumplimiento Semanal";
        //    chart.Title.Font.Bold = true;
        //    chart.Title.Font.Size = 14;
        //    chart.XAxis.Title.Text = "Semana";
        //    chart.XAxis.Title.Font.Bold = true;
        //    chart.YAxis.Title.Text = "Cumplimiento (%)";
        //    chart.YAxis.Title.Font.Bold = true;
        //    chart.YAxis.Format = "0%";
        //    chart.Legend.Remove();
        //    chart.Style = eChartStyle.Style1;
        //}
        private void GenerarGraficaEPPlus(ExcelWorksheet worksheet, DataTable datosGrafica, int filaInicio, int columnaInicio)
        {
            if (datosGrafica.Rows.Count == 0) return;

            var tempSheet = worksheet.Workbook.Worksheets.Add("TempData");
            tempSheet.Cells[1, 1].Value = "Semana";
            tempSheet.Cells[1, 2].Value = "Cumplimiento";
            tempSheet.Cells[1, 1].Style.Font.Bold = true;
            tempSheet.Cells[1, 2].Style.Font.Bold = true;

            for (int i = 0; i < datosGrafica.Rows.Count; i++)
            {
                tempSheet.Cells[i + 2, 1].Value = Convert.ToInt32(datosGrafica.Rows[i]["Semana"]);
                tempSheet.Cells[i + 2, 2].Value = Convert.ToDecimal(datosGrafica.Rows[i]["Total"]);
                tempSheet.Cells[i + 2, 2].Style.Numberformat.Format = "0%";
            }

            var chart = worksheet.Drawings.AddChart("CumplimientoChart", eChartType.ColumnClustered);
            chart.SetPosition(filaInicio, 0, columnaInicio, 0);
            chart.SetSize(600, 400);

            var serie = chart.Series.Add(tempSheet.Cells[2, 2, datosGrafica.Rows.Count + 1, 2],
                                         tempSheet.Cells[2, 1, datosGrafica.Rows.Count + 1, 1]) as ExcelBarChartSerie;

            if (serie != null)
            {
                serie.Header = "Cumplimiento Semanal";
                serie.DataLabel.ShowValue = true;

                // ==================== MODIFICAR COLOR DE LAS BARRAS ====================
                // Color azul corporativo (puedes cambiarlo al color que prefieras)
                serie.Fill.Style = eFillStyle.GradientFill;
                serie.Fill.Color = System.Drawing.Color.FromArgb(198, 224, 180);

                // Opcional: También puedes modificar el borde de las barras
                //serie.Border.Fill.Style = eFillStyle.SolidFill;
                //serie.Border.Fill.Color = System.Drawing.Color.FromArgb(41, 128, 185); // Azul pastel para el borde
                //serie.Border.Width = 1;
            }

            chart.Title.Text = "Cumplimiento Semanal";
            chart.Title.Font.Bold = true;
            chart.Title.Font.Size = 14;
            chart.XAxis.Title.Text = "Semana";
            chart.XAxis.Title.Font.Bold = true;
            chart.YAxis.Title.Text = "Cumplimiento (%)";
            chart.YAxis.Title.Font.Bold = true;
            chart.YAxis.Format = "0%";
            chart.Legend.Remove();
            chart.Style = eChartStyle.Style1;
        }

        #endregion

        #region Envío de correo

        public bool EnviarReportePorCorreo(string archivoExcel, List<int> semanasSeleccionadas)
        {
            try
            {
                if (string.IsNullOrEmpty(archivoExcel) || !File.Exists(archivoExcel))
                    throw new Exception("El archivo Excel no existe o no se pudo generar.");

                string semanasStr = string.Join(",", semanasSeleccionadas);
                string cuerpoHtml = $@"
                    <html>
                    <head>
                        <style>
                            body {{ font-family: Arial, sans-serif; color: #333; }}
                            .header {{ background-color: #2c3e50; color: white; padding: 20px; text-align: center; }}
                            .content {{ padding: 20px; }}
                            .footer {{ background-color: #f8f9fa; padding: 10px; text-align: center; font-size: 12px; color: #666; margin-top: 20px; }}
                        </style>
                    </head>
                    <body>
                        <div class='header'>
                            <h2>Reporte de Cumplimiento Semanal</h2>
                        </div>
                        <div class='content'>
                            <p>Estimado usuario,</p>
                            <p>Se adjunta el reporte general semanal de porcentajes de cumplimiento correspondiente a las semana(s) {semanasStr}.</p>
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

                using (MailMessage correo = new MailMessage())
                {
                    correo.From = new MailAddress(RemitenteEmail);

                    string[] destinatarios = DestinatariosEmail.Split(',');
                    foreach (var mail in destinatarios)
                    {
                        if (!string.IsNullOrWhiteSpace(mail))
                            correo.To.Add(mail.Trim());
                    }

                    correo.Subject = $"Reporte de Cumplimiento generado el {DateTime.Now:dd/MM/yyyy}";
                    correo.Body = cuerpoHtml;
                    correo.IsBodyHtml = true;

                    using (Attachment attachment = new Attachment(archivoExcel))
                    {
                        correo.Attachments.Add(attachment);

                        using (SmtpClient smtp = new SmtpClient(ServidorSMTP, PuertoSMTP))
                        {
                            smtp.Credentials = new NetworkCredential(RemitenteEmail, PasswordEmail);
                            smtp.EnableSsl = SSLCheck;
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

        public bool GenerarYEnviarReporte(List<int> semanasSeleccionadas)
        {
            string archivoExcel = null;
            try
            {
                archivoExcel = GenerarExcel(semanasSeleccionadas);
                bool enviado = EnviarReportePorCorreo(archivoExcel, semanasSeleccionadas);
                return enviado;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar y enviar el reporte: {ex.Message}");
            }
            finally
            {
                if (archivoExcel != null && File.Exists(archivoExcel))
                {
                    try { File.Delete(archivoExcel); } catch { }
                }
            }
        }

        #endregion
    }
}