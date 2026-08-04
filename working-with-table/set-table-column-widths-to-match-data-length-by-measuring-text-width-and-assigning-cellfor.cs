using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableColumnWidthExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with header and data rows.
            builder.StartTable();

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
            AddRow(builder, "Banana", "Ripe yellow banana", "$0.80");
            AddRow(builder, "Cherry", "Sweet dark cherries", "$3.50");
            AddRow(builder, "Watermelon", "Large green watermelon", "$7.00");

            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.FirstSection.Body.Tables[0];

            // Auto‑fit the columns to the contents of the cells.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            doc.Save("TableColumnWidths.docx");
        }

        // Helper method to add a data row to the table.
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
