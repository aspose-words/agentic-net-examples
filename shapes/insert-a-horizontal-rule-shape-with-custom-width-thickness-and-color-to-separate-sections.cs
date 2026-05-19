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

        // First section content.
        builder.Writeln("Section 1: Introduction");
        builder.Writeln("This is the first section of the document.");

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's appearance.
        HorizontalRuleFormat hrFormat = horizontalRule.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        hrFormat.WidthPercent = 80; // Width as 80% of the page width.
        hrFormat.Height = 5; // Thickness of the rule.
        hrFormat.Color = Color.DarkBlue; // Rule color.
        hrFormat.NoShade = true; // Solid color without 3‑D shading.

        // Second section content.
        builder.Writeln("Section 2: Details");
        builder.Writeln("This is the second section of the document.");

        // Save the document to disk.
        string outputFile = "HorizontalRule.docx";
        doc.Save(outputFile);

        // Verify that the file was created.
        if (!File.Exists(outputFile))
        {
            throw new InvalidOperationException($"The output file '{outputFile}' was not created.");
        }
    }
}
