using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Small 1x1 PNG image (transparent) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // ---------- First Row ----------
        // Cell (0,0) – insert an image.
        builder.InsertCell();
        using (var ms = new MemoryStream(imageBytes))
        {
            Shape image1 = builder.InsertImage(ms);
            image1.Width = 100;   // points
            image1.Height = 100;  // points
        }

        // Cell (0,1) – insert some text.
        builder.InsertCell();
        builder.Write("First image description");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cell (1,0) – insert text.
        builder.InsertCell();
        builder.Write("Second image description");

        // Cell (1,1) – insert another image.
        builder.InsertCell();
        using (var ms = new MemoryStream(imageBytes))
        {
            Shape image2 = builder.InsertImage(ms);
            image2.Width = 80;
            image2.Height = 80;
        }

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithImages.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
