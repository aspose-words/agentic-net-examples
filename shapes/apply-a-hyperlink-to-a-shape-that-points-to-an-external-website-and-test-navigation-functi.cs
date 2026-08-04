using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeHyperlinkExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 200, 100);
        // Apply a hyperlink that points to an external website.
        shape.HRef = "https://www.example.com/";
        // Open the link in a new window/tab.
        shape.Target = "New Window";
        // Optional tooltip shown when hovering over the shape.
        shape.ScreenTip = "Visit Example.com";

        // Define the output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ShapeWithHyperlink.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Validate that the shape contains the expected hyperlink.
        if (shape.HRef != "https://www.example.com/")
            throw new InvalidOperationException("The shape hyperlink was not set correctly.");
    }
}
