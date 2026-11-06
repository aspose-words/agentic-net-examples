using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Amount");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write((i * 10).ToString());
            builder.EndRow();
        }

        // Footer row with a field that sums the column above.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a formula field that calculates the sum of the cells above in this column.
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
