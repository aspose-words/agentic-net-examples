using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline cube shape with a size of 150x150 points.
        Shape shape = builder.InsertShape(ShapeType.Cube, 150, 150);
        shape.Name = "MyCube";

        // Set the alternative text for accessibility (screen readers, etc.).
        shape.AlternativeText = "Alt text for MyCube.";

        // Define the output file path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Shape_AltText.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform the user (no interactive prompts required).
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
