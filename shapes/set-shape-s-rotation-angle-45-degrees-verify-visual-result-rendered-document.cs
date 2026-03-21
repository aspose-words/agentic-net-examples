using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

class ShapeRotationDemo
{
    static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a new document and insert a shape (image) into it.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Base64-encoded PNG (1x1 pixel). Replace with any PNG data as needed.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        using (MemoryStream ms = new MemoryStream(pngBytes))
        {
            Shape shape = builder.InsertImage(ms);
            // Ensure the shape contains an image.
            if (!shape.CanHaveImage || !shape.HasImage)
                throw new InvalidOperationException("The inserted shape does not contain an image.");

            // -----------------------------------------------------------------
            // 2. Rotate the shape 45 degrees clockwise.
            // -----------------------------------------------------------------
            shape.Rotation = 45; // Positive value rotates clockwise.
        }

        // -----------------------------------------------------------------
        // 3. Save the document to disk.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "Shape.Rotate.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the saved document and render the rotated shape to an image.
        //    This visual output can be inspected to verify the rotation.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);
        if (loadedShape == null)
            throw new InvalidOperationException("Shape not found in the loaded document.");

        // Render the shape to a PNG file.
        string pngPath = Path.Combine(artifactsDir, "Shape.Rotate.png");
        ShapeRenderer renderer = new ShapeRenderer(loadedShape);
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png);
        renderer.Save(pngPath, imgOptions);

        // -----------------------------------------------------------------
        // 5. Inform the user that the process is complete.
        // -----------------------------------------------------------------
        Console.WriteLine("Document saved to: " + docPath);
        Console.WriteLine("Rendered shape image saved to: " + pngPath);
        Console.WriteLine("Open the PNG file to visually verify the 45° rotation.");
    }
}
