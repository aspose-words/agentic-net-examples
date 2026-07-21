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

        // A tiny red PNG image (2x2 pixels) encoded in Base64.
        // This avoids the need for System.Drawing and external image files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAADUlEQVQImWNgYGD4DwABBAEA7pV6WQAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image from the byte array as an inline shape.
        Shape shape = builder.InsertImage(imageBytes);

        // Lock the aspect ratio of the shape.
        shape.AspectRatioLocked = true;

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AspectRatioLocked.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
