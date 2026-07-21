using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableColumnWidthDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with varying text lengths.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Description");
            builder.InsertCell();
            builder.Write("Price");
            builder.EndRow();

            // Data rows.
            AddRow(builder, "Apple", "Fresh red apple", "$1.20");
            AddRow(builder, "Banana", "Ripe yellow banana from Ecuador", "$0.80");
            AddRow(builder, "Cherry", "Sweet cherries", "$3.50");
            AddRow(builder, "Dragonfruit", "Exotic tropical fruit with vibrant color", "$5.00");

            builder.EndTable();

            // Auto‑fit the table columns to the contents.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableColumnWidths.docx");
            doc.Save(outputPath);
        }

        // Helper method to insert a row with three cells.
        private static void AddRow(DocumentBuilder builder, string col1, string col2, string col3)
        {
            builder.InsertCell();
            builder.Write(col1);
            builder.InsertCell();
            builder.Write(col2);
            builder.InsertCell();
            builder.Write(col3);
            builder.EndRow();
        }
    }
}
