using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary directory for the sample files.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsSample");
        Directory.CreateDirectory(tempDir);

        // A 1x1 pixel PNG image encoded in Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6V6ZAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the document; this returns a Shape object.
        Shape shape = builder.InsertImage(imageBytes);

        // Lock the aspect ratio of the shape.
        shape.AspectRatioLocked = true;

        // Save the document.
        string outputPath = Path.Combine(tempDir, "AspectRatioLocked.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
