using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class ShapeHyperlinkExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image (or any shape) that will act as the clickable hyperlink.
        // Here we use InsertImage; you can replace the path with any valid image file.
        Shape shape = builder.InsertImage(@"C:\Images\Logo.jpg");

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

        // Save the document. The shape now behaves as a hyperlink when the document is opened in Word.
        doc.Save(@"C:\Output\ShapeWithHyperlink.docx");
    }
}
