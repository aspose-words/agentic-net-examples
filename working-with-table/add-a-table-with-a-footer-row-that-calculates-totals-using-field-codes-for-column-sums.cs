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

        // Begin the table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Sample data rows.
        string[,] data = {
            { "Apple",  "10", "0.5" },
            { "Banana", "5",  "0.3" },
            { "Orange", "8",  "0.4" }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            for (int j = 0; j < data.GetLength(1); j++)
            {
                builder.InsertCell();
                builder.Write(data[i, j]);
            }
            builder.EndRow();
        }

        // Footer row with totals calculated by field codes.
        builder.InsertCell();
        builder.Write("Total");
        // Quantity total.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", "0");
        // Price total.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Update fields so that totals are calculated before saving.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "TableWithFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
