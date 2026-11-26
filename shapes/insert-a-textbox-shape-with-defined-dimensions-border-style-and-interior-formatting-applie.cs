using System;
using System.Drawing;
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

        // Define dimensions for the textbox (in points).
        const double width = 200.0;   // 200 points (~2.78 inches)
        const double height = 100.0;  // 100 points (~1.39 inches)

        // Insert a textbox shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, width, height);

        // Configure border (stroke) style.
        textBox.StrokeColor = Color.DarkBlue;          // Border color.
        textBox.StrokeWeight = 2.0;                    // Border thickness (points).
        textBox.Stroke.DashStyle = DashStyle.Solid;    // Solid line.

        // Configure interior fill.
        textBox.FillColor = Color.LightYellow;        // Background color.

        // Add formatted text inside the textbox.
        // Move the builder cursor to the first paragraph of the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Center the paragraph horizontally.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        // Write the text.
        builder.Font.Size = 12;
        builder.Writeln("Hello, Aspose.Words!");

        // Save the document.
        string outputPath = "TextboxExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
