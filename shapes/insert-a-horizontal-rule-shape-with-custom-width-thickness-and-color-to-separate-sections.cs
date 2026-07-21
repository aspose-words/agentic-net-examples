using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content before the horizontal rule.
        builder.Writeln("Section 1: Introduction");
        builder.Writeln("This is the first section of the document.");

        // Insert the horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the rule's appearance.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center;
        format.WidthPercent = 80;      // 80% of the page width.
        format.Height = 5;             // Thickness in points.
        format.Color = Color.DarkBlue;
        format.NoShade = true;         // Solid color without 3‑D shading.

        // Add content after the rule.
        builder.Writeln("Section 2: Details");
        builder.Writeln("This is the second section of the document.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HorizontalRuleExample.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
