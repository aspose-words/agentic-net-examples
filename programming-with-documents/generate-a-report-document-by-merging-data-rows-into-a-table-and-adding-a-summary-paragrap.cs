using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsReport
{
    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            DataTable table = new DataTable("ReportData");
            table.Columns.Add("Product");
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("Price", typeof(decimal));

            table.Rows.Add("Apple", 10, 0.5m);
            table.Rows.Add("Banana", 5, 0.3m);
            table.Rows.Add("Carrot", 7, 0.2m);

            // Create a new blank document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a title
            builder.Writeln("Sales Report");
            builder.Writeln();

            // Start the table
            Table wordTable = builder.StartTable();

            // Header row
            builder.InsertCell();
            builder.Font.Bold = true;
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.InsertCell();
            builder.Write("Price");
            builder.EndRow();

            // Data rows
            builder.Font.Bold = false;
            foreach (DataRow row in table.Rows)
            {
                builder.InsertCell();
                builder.Write(row["Product"].ToString());

                builder.InsertCell();
                builder.Write(row["Quantity"].ToString());

                builder.InsertCell();
                builder.Write(string.Format("{0:C}", row["Price"]));
                builder.EndRow();
            }

            // End the table
            builder.EndTable();

            // Add a summary paragraph
            builder.Writeln();
            int totalRows = table.Rows.Count;
            decimal totalAmount = 0;
            foreach (DataRow row in table.Rows)
            {
                totalAmount += (int)row["Quantity"] * (decimal)row["Price"];
            }
            builder.Writeln($"Total items: {totalRows}");
            builder.Writeln($"Grand total: {totalAmount:C}");

            // Save the document
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            doc.Save(outputPath);
        }
    }
}
