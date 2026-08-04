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

        // Add some text before the horizontal rule.
        builder.Writeln("Section 1: Introduction");
        builder.Writeln("This is the first section of the document.");

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule: 80% width, 4 points thickness, blue color, solid (no shading).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.WidthPercent = 80;      // Width as a percentage of the page width.
        format.Height = 4;             // Thickness in points.
        format.Color = Color.Blue;    // Rule color.
        format.NoShade = true;         // Use solid color without 3D shading.

        // Add text after the horizontal rule.
        builder.Writeln("Section 2: Details");
        builder.Writeln("This is the second section after the horizontal rule.");

        // Save the document.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "HorizontalRuleExample.docx");
        doc.Save(outputFile);

        // Verify that the file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
