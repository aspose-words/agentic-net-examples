using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a tiny placeholder PNG (1x1 pixel) from a Base64 string.
        // This avoids the need for System.Drawing dependencies.
        string imagePath = Path.Combine(outputDir, "placeholder.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXK5wAAAABJRU5ErkJggg==");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the placeholder image as a shape.
        Shape shape = builder.InsertImage(imagePath);

        // Apply hyperlink to the shape.
        string url = "https://www.example.com/";
        shape.HRef = url;          // Destination URL.
        shape.Target = "New Window"; // Open in a new window.
        shape.ScreenTip = "Open example website";

        // Save the document.
        string docPath = Path.Combine(outputDir, "HyperlinkedShape.docx");
        doc.Save(docPath);

        // Validate that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception("Failed to save the document.");

        // Validate that the shape contains the expected hyperlink.
        if (shape.HRef != url)
            throw new Exception("Hyperlink was not set correctly on the shape.");

        Console.WriteLine("Document created successfully at: " + docPath);
    }
}
