using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Section 1");

        // Insert a horizontal rule shape with custom width, thickness, and color.
        // Width: 400 points, Thickness (height): 5 points.
        Shape hRule = builder.InsertShape(ShapeType.Line, 400, 5);
        hRule.Stroke.Color = Color.DarkBlue;   // Custom color.
        hRule.Stroke.Weight = 5;               // Thickness in points.

        // Add following text to demonstrate separation.
        builder.Writeln("Section 2");

        // Save the document.
        string outputPath = "HorizontalRuleExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
