using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data for the table.
        string[,] data = new string[,]
        {
            { "Item", "Quantity (kg)", "Price" },
            { "Apples", "20", "$1.20" },
            { "Bananas", "40", "$0.80" },
            { "Carrots", "50", "$0.60" }
        };

        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the table with the sample data.
        Table table = builder.StartTable();

        for (int row = 0; row < data.GetLength(0); row++)
        {
            for (int col = 0; col < data.GetLength(1); col++)
            {
                builder.InsertCell();
                builder.Write(data[row, col]);
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Let Aspose.Words automatically adjust column widths to fit the contents.
        // AutoFitToContents removes any preferred widths and recalculates the layout.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("TableWithMeasuredColumnWidths.docx");
    }
}
