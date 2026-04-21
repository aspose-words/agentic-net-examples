using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Security; // For SecurityElement.Escape
using Aspose.Words;
using Aspose.Words.Drawing;

public class MathMlAsShapesExample
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample MathML expressions.
        string[] mathMlExpressions = new[]
        {
            "<math><mi>a</mi><mo>+</mo><mi>b</mi></math>",
            "<math><mi>c</mi><mo>=</mo><mi>d</mi><mo>×</mo><mi>e</mi></math>",
            "<math><msup><mi>f</mi><mn>2</mn></msup><mo>-</mo><mi>g</mi></math>"
        };

        // Temporary folder for generated SVG files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "MathMlSvg");
        Directory.CreateDirectory(tempFolder);

        // Insert each MathML as an inline SVG image shape.
        for (int i = 0; i < mathMlExpressions.Length; i++)
        {
            string svgPath = Path.Combine(tempFolder, $"math{i}.svg");
            GenerateSimpleSvg(mathMlExpressions[i], svgPath);

            // Insert the SVG as an inline image.
            Shape shape = builder.InsertImage(svgPath);
            shape.Width = 60;   // points
            shape.Height = 20;  // points
            shape.WrapType = WrapType.Inline; // Ensure the shape is inline.

            // Add a space after the shape for readability.
            builder.Write(" ");
        }

        // End the paragraph.
        builder.Writeln();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MathMLShapes.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Clean up temporary SVG files.
        Directory.Delete(tempFolder, true);
    }

    // Generates a very simple SVG containing a readable approximation of the MathML.
    private static void GenerateSimpleSvg(string mathMl, string filePath)
    {
        // Remove XML tags to obtain a plain text representation.
        string readable = Regex.Replace(mathMl, "<.*?>", string.Empty).Trim();

        // Basic SVG template with the readable text.
        string svgContent = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<svg xmlns=""http://www.w3.org/2000/svg"" width=""200"" height=""50"">
  <text x=""0"" y=""35"" font-family=""Arial"" font-size=""24"">{SecurityElement.Escape(readable)}</text>
</svg>";

        File.WriteAllText(filePath, svgContent);
    }
}
