using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section text.
        builder.Writeln("Section 1");

        // Insert a horizontal line shape with custom width, thickness, and color.
        // Width: 400 points, Height (visual thickness): 5 points.
        Shape horizontalRule = builder.InsertShape(ShapeType.Line, 400, 5);
        // Set line color.
        horizontalRule.StrokeColor = Color.DarkSlateGray;
        // Set line thickness (weight) in points.
        horizontalRule.StrokeWeight = 5;
        // No text wrapping.
        horizontalRule.WrapType = WrapType.None;

        // Add a blank paragraph after the rule.
        builder.Writeln();

        // Second section text.
        builder.Writeln("Section 2");

        // Save the document.
        string outputPath = "HorizontalRule.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
