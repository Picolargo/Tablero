using Npgsql;
using NpgsqlTypes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Tablero
{
    /// <summary>
    /// Cumplimiento semanal por persona usando la MISMA regla del Reporte_Semanal_EPPlus:
    /// - Agrupar kilos por área y semana antes de calcular el porcentaje.
    /// - Para Despegue se usa Kg_prod_seco; para las otras siete áreas Kg_prod_term.
    /// - Descontar Kg_fuera_espec, dividir por Kg_meta y limitar al 100 %.
    /// - Si hay Despegue, ponderar Despegue 50 % y el promedio de las otras áreas 50 %.
    /// - Si no hay Despegue, el promedio de las otras áreas pesa 100 %.
    /// - Usar el número de semana del domingo de cierre (igual que el Excel original).
    /// </summary>
    public sealed class Reporte_Cumplimiento_Personal
    {
        private readonly string _connectionString;

        public Reporte_Cumplimiento_Personal(string connectionString)
        {
            _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        }

        public List<int> ObtenerAniosDisponibles()
        {
            var anios = new List<int>();
            using (var cn = new NpgsqlConnection(_connectionString))
            using (var cmd = new NpgsqlCommand(@"
                SELECT DISTINCT EXTRACT(YEAR FROM ""Fecha"")::integer AS anio
                FROM public.""Ficha""
                WHERE ""Fecha"" IS NOT NULL
                ORDER BY anio DESC", cn))
            {
                cn.Open();
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        anios.Add(reader.GetInt32(0));
            }

            if (anios.Count == 0)
                anios.Add(DateTime.Today.Year);
            return anios;
        }

        public DataTable Consultar(string nivel, int anio, IEnumerable<int> semanas)
        {
            if (nivel != "Supervisor" && nivel != "Jefe de Turno")
                throw new ArgumentException("El nivel debe ser Supervisor o Jefe de Turno.", nameof(nivel));
            if (anio < 2000 || anio > 2200)
                throw new ArgumentOutOfRangeException(nameof(anio));
            if (semanas == null)
                throw new ArgumentNullException(nameof(semanas));

            int[] seleccionadas = semanas.Distinct().OrderBy(s => s).ToArray();
            if (seleccionadas.Length == 0 || seleccionadas.Any(s => s < 1 || s > 53))
                throw new ArgumentException("Selecciona al menos una semana válida (1 a 53).", nameof(semanas));

            // La relación cambia con el rol; NO se concatena ningún valor proporcionado por el usuario.
            string columnaPersona = nivel == "Supervisor" ? @"f.""ID_user""" : @"f.""ID_Jefe""";

            string sql = $@"
WITH usuarios AS (
    SELECT ""ID_User"", ""Usuario"", ""No_Empleado""
    FROM public.""Usuarios""
    WHERE ""Nivel"" = @nivel AND ""Activo"" = TRUE
),
semanas AS (
    SELECT unnest(@semanas) AS semana
),
por_area AS (
    SELECT
        u.""ID_User"" AS id_persona,
        EXTRACT(WEEK FROM DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days')::integer AS semana,
        f.""Area"" AS area,
        COUNT(*)::integer AS fichas,
        SUM(COALESCE(f.""Kg_meta"", 0)) AS meta,
        SUM(CASE WHEN f.""Area"" = 'Despegue'
                 THEN COALESCE(f.""Kg_prod_seco"", 0)
                 ELSE COALESCE(f.""Kg_prod_term"", 0) END) AS producido,
        SUM(COALESCE(f.""Kg_fuera_espec"", 0)) AS fuera_espec
    FROM public.""Ficha"" f
    INNER JOIN usuarios u ON {columnaPersona} = u.""ID_User""
    WHERE EXTRACT(YEAR FROM f.""Fecha"")::integer = @anio
      AND EXTRACT(WEEK FROM DATE_TRUNC('week', f.""Fecha"") + INTERVAL '6 days')::integer = ANY(@semanas)
      AND f.""Area"" IN ('Despegue', 'Evaporado', 'Grind', 'Inspeccion',
                           'Polvos', 'Empacado', 'Revolturas', 'Máquinas')
    GROUP BY u.""ID_User"", semana, f.""Area""
),
cumplimiento_area AS (
    SELECT id_persona, semana, area, fichas,
           CASE WHEN meta > 0 THEN
               LEAST(ROUND(((producido - fuera_espec) / meta) * 100, 2), 100)
           ELSE 0 END AS cumplimiento
    FROM por_area
),
resumen_persona AS (
    SELECT id_persona, semana,
           MAX(cumplimiento) FILTER (WHERE area = 'Despegue') AS despegue,
           AVG(cumplimiento) FILTER (WHERE area <> 'Despegue') AS otras_areas,
           SUM(fichas)::integer AS fichas,
           COUNT(*)::integer AS areas
    FROM cumplimiento_area
    GROUP BY id_persona, semana
),
ponderado AS (
    SELECT u.""ID_User"" AS id, u.""Usuario"" AS nombre,
           u.""No_Empleado"" AS no_empleado, s.semana,
           r.fichas, r.areas,
           CASE WHEN r.despegue IS NOT NULL
                THEN ROUND(r.despegue * 0.5, 2) END AS despegue_ponderado,
           CASE WHEN r.otras_areas IS NOT NULL THEN
                ROUND(r.otras_areas *
                      (CASE WHEN r.despegue IS NOT NULL THEN 0.5 ELSE 1 END), 2)
           END AS otras_ponderadas,
           CASE WHEN r.areas IS NOT NULL THEN
                ROUND(COALESCE(r.despegue * 0.5, 0) +
                      COALESCE(r.otras_areas *
                        (CASE WHEN r.despegue IS NOT NULL THEN 0.5 ELSE 1 END), 0), 2)
           END AS total
    FROM usuarios u
    CROSS JOIN semanas s
    LEFT JOIN resumen_persona r ON r.id_persona = u.""ID_User"" AND r.semana = s.semana
)
SELECT id AS ""ID"", @anio AS ""Año"", semana AS ""Semana"",
       nombre AS ""Nombre"", no_empleado AS ""No. empleado"",
       despegue_ponderado AS ""Despegue (%)"",
       otras_ponderadas AS ""Demás áreas (%)"",
       total AS ""Total (%)"",
       COALESCE(fichas, 0) AS ""Fichas"",
       COALESCE(areas, 0) AS ""Áreas""
FROM ponderado
ORDER BY nombre, id, semana;";

            var resultado = new DataTable();
            using (var cn = new NpgsqlConnection(_connectionString))
            using (var cmd = new NpgsqlCommand(sql, cn))
            {
                cmd.CommandTimeout = 120;
                cmd.Parameters.Add("anio", NpgsqlDbType.Integer).Value = anio;
                cmd.Parameters.Add("nivel", NpgsqlDbType.Varchar).Value = nivel;
                cmd.Parameters.Add("semanas", NpgsqlDbType.Array | NpgsqlDbType.Integer).Value = seleccionadas;
                using (var adaptador = new NpgsqlDataAdapter(cmd))
                    adaptador.Fill(resultado);
            }

            // Igual que CalcularPromedioMensual de EPPlus: promediar únicamente
            // los totales de las semanas que tienen valor mayor que cero.
            resultado.Columns.Add("Promedio período (%)", typeof(decimal));
            var promediosPorUsuario = resultado.AsEnumerable()
                .GroupBy(r => Convert.ToInt32(r["ID"]))
                .ToDictionary(
                    g => g.Key,
                    g => g.Where(r => r["Total (%)"] != DBNull.Value &&
                                     Convert.ToDecimal(r["Total (%)"]) > 0)
                          .Select(r => Convert.ToDecimal(r["Total (%)"]))
                          .ToList());

            foreach (DataRow fila in resultado.Rows)
            {
                List<decimal> valores = promediosPorUsuario[Convert.ToInt32(fila["ID"])];
                fila["Promedio período (%)"] = valores.Count > 0
                    ? (object)Math.Round(valores.Average(), 2, MidpointRounding.AwayFromZero)
                    : DBNull.Value;
            }

            // Solo cambia la posición visual del promedio; la columna ID se oculta en el Grid.
            resultado.Columns["Promedio período (%)"].SetOrdinal(resultado.Columns["Total (%)"].Ordinal + 1);
            return resultado;
        }
    }
}
