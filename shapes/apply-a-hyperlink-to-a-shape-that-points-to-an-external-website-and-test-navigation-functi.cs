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

        // Insert a simple rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);

        // Apply a hyperlink that points to an external website.
        const string url = "https://www.example.com";
        shape.HRef = url;               // Set the hyperlink address.
        shape.Target = "New Window";    // Open the link in a new browser window.
        shape.ScreenTip = "Open Example.com"; // Tooltip shown on hover.

        // Validate that the hyperlink was set correctly.
        if (shape.HRef != url)
            throw new InvalidOperationException("Failed to set the shape hyperlink.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyperlinkedShape.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
