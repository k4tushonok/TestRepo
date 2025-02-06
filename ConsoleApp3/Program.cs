using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using ExcelDataReader;
using System.Globalization;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string excelFilePath = @"C:\Users\ea\Desktop\VENTES.xlsx";
        string exportParetoPath = @"C:\Users\ea\Desktop\ParetoChart.xlsx";
        string exportTopProductsPath = @"C:\Users\ea\Desktop\TopProducts.xlsx";
        string connectionString = "Server=(local)\\SQLEXPRESS;Database=Test;Trusted_Connection=True;TrustServerCertificate=True;";

        EnsureTableExists(connectionString);
        ImportExcelToSql(excelFilePath, connectionString);
        ExportParetoChart(connectionString, exportParetoPath);
        ExportTopProducts(connectionString, exportTopProductsPath, 0.8);
    }

    static void EnsureTableExists(string connectionString)
    {
        string createTableQuery = @"
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'VENTES')
            BEGIN
                CREATE TABLE [dbo].[VENTES] (
                    [CLI_ID] int NOT NULL, 
                    [VNT_DATE] smalldatetime NOT NULL, 
                    [PRD_ID] int NOT NULL, 
                    [VNT_COUNT] smallint NOT NULL, 
                    [VNT_PRICE] decimal(8,2) NOT NULL
                );
            END";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            using (SqlCommand command = new SqlCommand(createTableQuery, connection))
            {
                command.ExecuteNonQuery();
            }
        }
    }

    static void ImportExcelToSql(string filePath, string connectionString)
    {
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var dataSet = reader.AsDataSet();
            var table = dataSet.Tables[0];

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                for (int i = 1; i < table.Rows.Count; i++)
                {
                    if (!int.TryParse(table.Rows[i][0].ToString(), out int cliId) ||
                        !DateTime.TryParseExact(table.Rows[i][1].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime vntDate) ||
                        !int.TryParse(table.Rows[i][2].ToString(), out int prdId) ||
                        !short.TryParse(table.Rows[i][3].ToString(), out short vntCount) ||
                        !decimal.TryParse(table.Rows[i][4].ToString(), out decimal vntPrice))
                    {
                        continue;
                    }

                    string query = @"
                        MERGE INTO VENTES AS target
                        USING (SELECT @CLI_ID AS CLI_ID, @VNT_DATE AS VNT_DATE, @PRD_ID AS PRD_ID, @VNT_COUNT AS VNT_COUNT, @VNT_PRICE AS VNT_PRICE) AS source
                        ON target.CLI_ID = source.CLI_ID AND target.VNT_DATE = source.VNT_DATE AND target.PRD_ID = source.PRD_ID
                        WHEN NOT MATCHED THEN 
                        INSERT (CLI_ID, VNT_DATE, PRD_ID, VNT_COUNT, VNT_PRICE)
                        VALUES (source.CLI_ID, source.VNT_DATE, source.PRD_ID, source.VNT_COUNT, source.VNT_PRICE);";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@CLI_ID", cliId);
                        command.Parameters.AddWithValue("@VNT_DATE", vntDate);
                        command.Parameters.AddWithValue("@PRD_ID", prdId);
                        command.Parameters.AddWithValue("@VNT_COUNT", vntCount);
                        command.Parameters.AddWithValue("@VNT_PRICE", vntPrice);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }

    static void ExportParetoChart(string connectionString, string exportPath)
    {
        string query = @"
            WITH SalesData AS (
                SELECT 
                    PRD_ID,
                    SUM(VNT_COUNT) AS TotalCount,
                    SUM(VNT_COUNT * VNT_PRICE) AS TotalSales
                FROM VENTES
                GROUP BY PRD_ID
            ),
            OrderedSales AS (
                SELECT 
                    PRD_ID, 
                    TotalCount, 
                    TotalSales,
                    NTILE(20) OVER (ORDER BY TotalSales DESC) AS GroupNumber
                FROM SalesData
            ),
            GroupedData AS (
                SELECT 
                    GroupNumber,
                    SUM(TotalCount) AS GroupCount,
                    SUM(TotalSales) AS GroupSales
                FROM OrderedSales
                GROUP BY GroupNumber
            ),
            TotalValues AS (
                SELECT 
                    SUM(TotalCount) AS TotalCountAll, 
                    SUM(TotalSales) AS TotalSalesAll
                FROM SalesData
            ),
            Cumulative AS (
                SELECT 
                    GroupNumber,
                    GroupCount,
                    GroupSales,
                    SUM(GroupCount) OVER (ORDER BY GroupNumber) * 100.0 / (SELECT TotalCountAll FROM TotalValues) AS CumulativePercentCount,
                    SUM(GroupSales) OVER (ORDER BY GroupNumber) * 100.0 / (SELECT TotalSalesAll FROM TotalValues) AS CumulativePercentSales
                FROM GroupedData
            )
            SELECT 
                GroupNumber AS [Step],
                ROUND(CumulativePercentCount, 2) AS PercentCount,
                ROUND(CumulativePercentSales, 2) AS PercentSales
            FROM Cumulative
            ORDER BY GroupNumber;";

        using (SqlConnection connection = new SqlConnection(connectionString))
        using (SqlCommand command = new SqlCommand(query, connection))
        {
            connection.Open();
            using (SqlDataReader reader = command.ExecuteReader())
            {
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("ParetoChart");

                worksheet.Cell(1, 1).Value = "Step";
                worksheet.Cell(1, 2).Value = "Cumulative % Count";
                worksheet.Cell(1, 3).Value = "Cumulative % Sales";

                int row = 2;
                while (reader.Read())
                {
                    worksheet.Cell(row, 1).Value = reader.GetInt64(0);
                    worksheet.Cell(row, 2).Value = reader.GetDecimal(1);
                    worksheet.Cell(row, 3).Value = reader.GetDecimal(2);
                    row++;
                }

                workbook.SaveAs(exportPath);
            }
        }
    }

    static void ExportTopProducts(string connectionString, string exportPath, double targetPercent)
    {
        string query = @"
            WITH SalesData AS (
                SELECT 
                    PRD_ID,
                    SUM(VNT_COUNT * VNT_PRICE) AS TotalSales
                FROM VENTES
                GROUP BY PRD_ID
            ),
            RankedSales AS (
                SELECT 
                    PRD_ID,
                    TotalSales,
                    ROW_NUMBER() OVER (ORDER BY TotalSales DESC) AS RN,
                    SUM(TotalSales) OVER (ORDER BY TotalSales DESC) AS CumulativeSales,
                    SUM(TotalSales) OVER () AS TotalSalesAll
                FROM SalesData
            ),
            Threshold AS (
                SELECT MIN(RN) AS ThresholdRN
                FROM RankedSales
                WHERE CumulativeSales * 1.0 / TotalSalesAll >= @targetPercent
            )
            SELECT 
                RS.PRD_ID,
                ROUND(RS.CumulativeSales * 100.0 / RS.TotalSalesAll, 2) AS SalesShare,
                RS.TotalSales
            FROM RankedSales RS
            CROSS JOIN Threshold T
            WHERE RS.RN <= T.ThresholdRN
            ORDER BY RS.TotalSales DESC;";

        using (SqlConnection connection = new SqlConnection(connectionString))
        using (SqlCommand command = new SqlCommand(query, connection))
        {
            command.Parameters.AddWithValue("@targetPercent", targetPercent);

            connection.Open();
            using (SqlDataReader reader = command.ExecuteReader())
            {
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("TopProducts");

                worksheet.Cell(1, 1).Value = "PRD_ID";
                worksheet.Cell(1, 2).Value = "Sales Share (%)";
                worksheet.Cell(1, 3).Value = "Total Sales";

                int row = 2;
                while (reader.Read())
                {
                    worksheet.Cell(row, 1).Value = reader.GetInt32(0);
                    worksheet.Cell(row, 2).Value = reader.GetDecimal(1);
                    worksheet.Cell(row, 3).Value = reader.GetDecimal(2);
                    row++;
                }

                workbook.SaveAs(exportPath);
            }
        }
    }
}



