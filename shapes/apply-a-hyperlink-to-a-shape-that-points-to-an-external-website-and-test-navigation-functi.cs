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

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        shape.WrapType = WrapType.None;                     // Make it floating.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 100;                                   // Position from the left of the page.
        shape.Top = 100;                                    // Position from the top of the page.

        // Apply hyperlink properties.
        const string url = "https://www.example.com/";
        shape.HRef = url;                                   // Destination URL.
        shape.Target = "New Window";                        // Open in a new browser window.
        shape.ScreenTip = "Open Example.com";               // Tooltip shown on hover.

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeWithHyperlink.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Load the document again and verify the hyperlink on the shape.
        Document loadedDoc = new Document(outputPath);
        Shape loadedShape = (Shape)loadedDoc.GetChildNodes(NodeType.Shape, true)[0];
        if (loadedShape.HRef != url)
            throw new InvalidOperationException("The shape hyperlink was not set correctly.");

        // The example finishes without requiring user interaction.
    }
}
