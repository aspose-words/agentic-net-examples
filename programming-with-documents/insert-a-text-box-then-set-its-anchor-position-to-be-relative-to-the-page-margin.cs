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

        // Insert a floating text box shape.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Make the shape floating so we can set its position.
        textBoxShape.WrapType = WrapType.None;

        // Anchor the shape relative to the page margins.
        textBoxShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        textBoxShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Optional: set an offset from the margins.
        textBoxShape.Left = 50; // 50 points from the left margin.
        textBoxShape.Top = 50;  // 50 points from the top margin.

        // Add text inside the text box.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Write("Hello from a text box anchored to the page margin.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TextBoxAnchorMargin.docx");
        doc.Save(outputPath);
    }
}
