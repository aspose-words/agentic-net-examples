using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Apply a thick top border (4 points) and a thin bottom border (1 point).
        // The last parameter 'true' overrides any existing cell borders.
        table.SetBorder(BorderType.Top, LineStyle.Single, 4.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);

        // Add a simple 2‑cell row to the table.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define the output file name.
        string outputPath = "TableBorders.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
