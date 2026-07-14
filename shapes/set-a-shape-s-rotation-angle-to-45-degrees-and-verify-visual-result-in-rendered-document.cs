using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "RotatedShape.docx");

        // A 1x1 red PNG image encoded in Base64.
        // This avoids the need for external image files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8z8BQDwAF/AL+XK6XAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image from the byte array.
        Shape shape = builder.InsertImage(imageBytes);

        // Rotate the shape 45 degrees clockwise.
        shape.Rotation = 45;

        // Save the document.
        doc.Save(docPath);

        // Validation: ensure the file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Validation: ensure the shape's rotation property is set correctly.
        if (Math.Abs(shape.Rotation - 45) > 0.001)
            throw new InvalidOperationException($"Shape rotation is {shape.Rotation}, expected 45.");
    }
}
