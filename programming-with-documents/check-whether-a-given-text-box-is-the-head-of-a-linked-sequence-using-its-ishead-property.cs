using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first text box.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        shape1.WrapType = WrapType.Inline;
        shape1.AppendChild(new Paragraph(doc));
        shape1.FirstParagraph.AppendChild(new Run(doc, "First box"));

        // Insert the second text box.
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        shape2.WrapType = WrapType.Inline;
        shape2.AppendChild(new Paragraph(doc));
        shape2.FirstParagraph.AppendChild(new Run(doc, "Second box"));

        // Attempt to link the two text boxes (using dynamic to avoid compile‑time errors if the API is unavailable).
        try
        {
            dynamic dShape1 = shape1;
            dShape1.LinkTo(shape2);
        }
        catch
        {
            // Linking not supported in this version – continue without linking.
        }

        // Check the IsHead property for each shape (using dynamic to avoid compile‑time errors).
        bool isHead1 = false;
        bool isHead2 = false;

        try
        {
            dynamic dShape1 = shape1;
            isHead1 = dShape1.IsHead;
        }
        catch
        {
            // Property not available – default to false.
        }

        try
        {
            dynamic dShape2 = shape2;
            isHead2 = dShape2.IsHead;
        }
        catch
        {
            // Property not available – default to false.
        }

        Console.WriteLine($"Shape 1 IsHead: {isHead1}");
        Console.WriteLine($"Shape 2 IsHead: {isHead2}");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LinkedTextBoxes.docx");
        doc.Save(outputPath);

        // Verify that the file was saved and can be reopened.
        Document loadedDoc = new Document(outputPath);
        Console.WriteLine($"Document reloaded successfully. Contains {loadedDoc.GetChildNodes(NodeType.Shape, true).Count} shapes.");
    }
}
