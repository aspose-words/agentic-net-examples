using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace HorizontalRuleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("Section 1: Introduction");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

            // Insert a horizontal rule shape with custom width, thickness, and color.
            // Use ShapeType.Line because ShapeType.HorizontalLine does not exist.
            // Width = 400 points, Height (thickness) = 5 points.
            Shape horizontalRule = builder.InsertShape(ShapeType.Line, 400, 5);
            horizontalRule.StrokeColor = Color.Blue;   // Set line color.
            horizontalRule.StrokeWeight = 5;           // Set line thickness (points).

            // Add some text after the rule to demonstrate separation.
            builder.Writeln();
            builder.Writeln("Section 2: Details");
            builder.Writeln("Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Save the document.
            string outputPath = "HorizontalRule.docx";
            doc.Save(outputPath);

            // Validate that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
