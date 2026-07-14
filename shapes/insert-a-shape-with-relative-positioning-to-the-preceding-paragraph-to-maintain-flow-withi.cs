using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add the first paragraph of text.
        builder.Writeln("This is the first paragraph. The shape will be positioned relative to this paragraph.");

        // Insert a floating rectangle shape.
        // Position it relative to the preceding paragraph (RelativeVerticalPosition.Paragraph)
        // and to the left margin (RelativeHorizontalPosition.Margin).
        // Left and top offsets are set to 0 to align with the paragraph start.
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Margin, 0,
            RelativeVerticalPosition.Paragraph, 0,
            100,   // width in points
            50,    // height in points
            WrapType.Square);

        // Optional formatting for visibility.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;

        // Continue with more text after the shape.
        builder.Writeln("This is the second paragraph, appearing after the shape.");

        // Save the document to the local file system.
        string outputPath = "ShapeRelativeToParagraph.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");

        // Inform the user (optional, not required for non‑interactive execution).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
