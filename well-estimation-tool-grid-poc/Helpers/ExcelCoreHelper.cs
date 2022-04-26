using OfficeOpenXml;
using System.Data.SqlClient;

namespace well_estimation_tool_grid_poc.Helpers
{
    public class ExcelCoreHelper
    {
        //public async Task MainExcelMethod(string[] args)
        //{
        //    try
        //    {
        //        string s = null;
        //        var d = new DirectoryInfo(@"C:\Test");
        //        var files = d.GetFiles("*.xlsx");
        //        var usersList = new List<User>();

        //        foreach (var file in files)
        //        {
        //            var fileName = file.FullName;
        //            using var package = new ExcelPackage(file);
        //            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //            var currentSheet = package.Workbook.Worksheets;
        //            var workSheet = currentSheet.First();
        //            var noOfCol = workSheet.Dimension.End.Column;
        //            var noOfRow = workSheet.Dimension.End.Row;
        //            for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
        //            {
        //                var user = new User
        //                {
        //                    GameCode = workSheet.Cells[rowIterator, 1].Value?.ToString(),
        //                    Count = Convert.ToInt32(workSheet.Cells[rowIterator, 2].Value),
        //                    Email = workSheet.Cells[rowIterator, 3].Value?.ToString(),
        //                    Status = Convert.ToInt32(workSheet.Cells[rowIterator, 4].Value)
        //                };
        //                usersList.Add(user);
        //            }
        //        }
        //        var conn = ConfigurationManager.ConnectionStrings["Development"].ConnectionString;
        //        await using var connString = new SqlConnection(conn);
        //        connString.Open();
        //        await BulkWriter.InsertAsync(usersList, "[Orders]", connString, CancellationToken.None);
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e);
        //        throw;
        //    }

        //}
        //public class BulkWriter
        //{
        //    private static readonly ConcurrentDictionary<Type, SqlBulkCopyColumnMapping[]> ColumnMapping =
        //        new ConcurrentDictionary<Type, SqlBulkCopyColumnMapping[]>();
        //    public static async Task InsertAsync<T>(IEnumerable<T> items, string tableName, SqlConnection connection,
        //        CancellationToken cancellationToken)
        //    {
        //        using var bulk = new SqlBulkCopy(connection);
        //        await using var reader = ObjectReader.Create(items);
        //        bulk.DestinationTableName = tableName;
        //        foreach (var colMap in GetColumnMappings<T>())
        //            bulk.ColumnMappings.Add(colMap);
        //        await bulk.WriteToServerAsync(reader, cancellationToken);
        //    }
        //    private static IEnumerable<SqlBulkCopyColumnMapping> GetColumnMappings<T>() =>
        //        ColumnMapping.GetOrAdd(typeof(T),
        //            type =>
        //                type.GetProperties()
        //                    .Select(p => new SqlBulkCopyColumnMapping(p.Name, p.Name)).ToArray());
        //}
    }
}
