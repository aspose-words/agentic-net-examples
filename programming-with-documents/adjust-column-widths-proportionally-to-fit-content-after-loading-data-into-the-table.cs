using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AdjustTableColumns
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Description");
            builder.InsertCell();
            builder.Write("Price");
            builder.EndRow();

            // Add data rows.
            AddRow(builder, "Apple", "Fresh red apples from the orchard", "$1.20");
            AddRow(builder, "Banana", "Ripe bananas, sweet and soft", "$0.80");
            AddRow(builder, "Cherry", "Organic cherries, packed in a box", "$3.50");
            AddRow(builder, "Watermelon", "Large watermelon, perfect for summer picnics", "$5.00");

            // End the table.
            builder.EndTable();

            // Adjust column widths proportionally to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            doc.Save("AdjustedTable.docx");
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
