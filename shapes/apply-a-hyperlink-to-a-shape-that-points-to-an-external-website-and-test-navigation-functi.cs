using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class HyperlinkShapeExample
{
    public static void Main()
    {
        // Define output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HyperlinkShape.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 150, 80);
        // Apply hyperlink properties.
        shape.HRef = "https://www.example.com/";
        shape.Target = "New Window";
        shape.ScreenTip = "Open Example website";

        // Save the document.
        doc.Save(outputPath);

        // Load the document back to verify the hyperlink.
        Document loadedDoc = new Document(outputPath);
        Shape loadedShape = (Shape)loadedDoc.GetChildNodes(NodeType.Shape, true)[0];

        // Validate that the hyperlink address matches the expected value.
        if (loadedShape.HRef != "https://www.example.com/")
            throw new InvalidOperationException("Hyperlink address was not set correctly.");

        // Optionally, validate other properties.
        if (loadedShape.Target != "New Window")
            throw new InvalidOperationException("Hyperlink target was not set correctly.");

        if (loadedShape.ScreenTip != "Open Example website")
            throw new InvalidOperationException("Hyperlink screen tip was not set correctly.");

        // Indicate successful completion.
        Console.WriteLine("Document created and hyperlink verified successfully.");
    }
}
