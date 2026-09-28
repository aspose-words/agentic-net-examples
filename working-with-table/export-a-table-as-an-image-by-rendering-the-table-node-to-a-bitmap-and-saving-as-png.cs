using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Define the output image path.
        string imagePath = "Table.png";

        // Render only the first page (which contains the table) to a PNG image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Use PageSet to specify which pages to render.
            PageSet = new PageSet(0) // Render page index 0 only.
        };

        // Save the document as an image.
        doc.Save(imagePath, options);

        // Verify that the image file was created.
        if (!File.Exists(imagePath))
        {
            throw new InvalidOperationException($"Failed to create image file: {imagePath}");
        }
    }
}
