using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table construction.
        builder.EndTable();

        // Set the table's vertical alignment on the page to center.
        // For floating tables the vertical alignment is controlled by RelativeVerticalAlignment.
        table.RelativeVerticalAlignment = VerticalAlignment.Center;

        // Save the document to the local file system.
        string outputPath = "TableVerticalAlignment.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }
    }
}
