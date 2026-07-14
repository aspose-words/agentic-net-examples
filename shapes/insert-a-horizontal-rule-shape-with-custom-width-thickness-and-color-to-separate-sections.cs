using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some content before the horizontal rule.
        builder.Writeln("Section 1: Introduction");
        builder.Writeln("This is the first section of the document.");

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule: 80% width, 5 points thickness, blue color, solid (no shading).
        HorizontalRuleFormat hrFormat = horizontalRule.HorizontalRuleFormat;
        hrFormat.WidthPercent = 80;   // Length as a percentage of the page width.
        hrFormat.Height = 5;          // Thickness in points.
        hrFormat.Color = Color.Blue; // Rule color.
        hrFormat.NoShade = true;      // Use solid color without 3‑D shading.

        // Write content after the horizontal rule.
        builder.Writeln("Section 2: Details");
        builder.Writeln("This is the second section of the document.");

        // Save the document.
        string outputFile = "HorizontalRuleExample.docx";
        doc.Save(outputFile);

        // Validate that the file was created.
        if (!File.Exists(outputFile))
        {
            throw new Exception($"Output file was not created: {outputFile}");
        }
    }
}
