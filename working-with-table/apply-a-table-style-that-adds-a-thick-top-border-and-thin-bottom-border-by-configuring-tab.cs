using System;
using System.Drawing;
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

        builder.EndTable();

        // Apply a thick top border (3 points) and a thin bottom border (1 point) to the whole table.
        // The last parameter 'true' overrides any existing cell borders.
        table.SetBorder(BorderType.Top, LineStyle.Single, 3.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableBorders.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
