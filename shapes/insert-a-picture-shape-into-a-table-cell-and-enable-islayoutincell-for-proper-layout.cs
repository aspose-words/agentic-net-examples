using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with two cells.
        builder.StartTable();
        builder.InsertCell(); // First (empty) cell.
        builder.InsertCell(); // Second cell – the image will be placed here.

        // Insert a simple placeholder PNG image (1x1 pixel) from a byte array.
        // This avoids the need for System.Drawing dependencies.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        Shape imageShape = builder.InsertImage(pngBytes);

        // Make the shape floating so IsLayoutInCell takes effect.
        imageShape.WrapType = WrapType.None;
        // Enable layout inside the table cell.
        imageShape.IsLayoutInCell = true;
        // Set explicit size (optional).
        imageShape.Width = 80;
        imageShape.Height = 80;

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "TableImageShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        // Validate that the shape has IsLayoutInCell set.
        Shape foundShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (foundShape == null || !foundShape.IsLayoutInCell)
            throw new Exception("The image shape does not have IsLayoutInCell enabled.");
    }
}
