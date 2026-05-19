using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x1 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Apply a solid light‑blue shading to the whole table.
        table.SetShading(TextureIndex.TextureSolid, Color.LightBlue, Color.Empty);

        // Define output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableShading.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        // Optionally, inform that the process completed (no interactive input required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
