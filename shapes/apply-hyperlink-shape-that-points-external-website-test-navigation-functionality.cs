using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class ShapeHyperlinkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape that will act as the clickable hyperlink.
        // Width and height are in points (1 point = 1/72 inch).
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;
        shape.WrapType = WrapType.Inline;

        // Apply hyperlink properties to the shape.
        shape.HRef = "https://www.example.com/";   // Destination URL.
        shape.Target = "New Window";               // Open in a new browser window.
        shape.ScreenTip = "Open Example.com";      // Tooltip shown on mouse hover.

        // Optionally, verify that the properties are set correctly (useful for automated tests).
        if (shape.HRef != "https://www.example.com/")
            throw new InvalidOperationException("HRef was not set correctly.");
        if (shape.Target != "New Window")
            throw new InvalidOperationException("Target was not set correctly.");
        if (shape.ScreenTip != "Open Example.com")
            throw new InvalidOperationException("ScreenTip was not set correctly.");

        // Save the document in the current directory.
        doc.Save("ShapeWithHyperlink.docx");
    }
}
