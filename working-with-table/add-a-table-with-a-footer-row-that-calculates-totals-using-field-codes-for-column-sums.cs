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

        // Add a few data rows.
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");

            builder.InsertCell();
            builder.Write((i * 2).ToString()); // Quantity

            builder.InsertCell();
            builder.Write((i * 10).ToString()); // Price

            builder.EndRow();
        }

        // Footer row with field codes that sum the columns above.
        builder.InsertCell();
        builder.Write("Total");

        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null); // Sum of Quantity column.

        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null); // Sum of Price column.

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "TableWithFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
