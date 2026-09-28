using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string folder = Path.Combine(Path.GetTempPath(), "AsposeShapesExample");
        Directory.CreateDirectory(folder);

        // Create a simple 1x1 PNG image (transparent) from a Base64 string.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(folder, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with two cells.
        builder.StartTable();

        // First cell – just some text.
        builder.InsertCell();
        builder.Writeln("First cell");

        // Second cell – insert the picture shape.
        builder.InsertCell();

        // Insert the image; this returns a Shape object.
        Shape pictureShape = builder.InsertImage(imagePath);
        // Enable layout inside the table cell.
        pictureShape.IsLayoutInCell = true;

        // Optionally adjust size.
        pictureShape.Width = 80;
        pictureShape.Height = 80;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(folder, "output.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        // Validate that the shape has IsLayoutInCell set.
        Shape savedShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        if (!savedShape.IsLayoutInCell)
            throw new Exception("IsLayoutInCell property was not set on the picture shape.");

        // Indicate success.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
