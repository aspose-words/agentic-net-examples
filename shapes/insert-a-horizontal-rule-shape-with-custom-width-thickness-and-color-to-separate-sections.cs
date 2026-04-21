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

        // Add some text before the horizontal rule.
        builder.Writeln("Section 1: Introduction");
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule: width, thickness, and color.
        HorizontalRuleFormat hrFormat = horizontalRule.HorizontalRuleFormat;
        hrFormat.WidthPercent = 80;          // Width as a percentage of the page width.
        hrFormat.Height = 5;                 // Thickness in points.
        hrFormat.Color = Color.DarkGray;    // Solid color.
        hrFormat.NoShade = true;            // Disable 3‑D shading.

        // Add text after the horizontal rule.
        builder.Writeln("Section 2: Details");
        builder.Writeln("Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HorizontalRule.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The output file was not created: {outputPath}");
    }
}
