using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;

public class MathMlShapesExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the inline MathML equations.
        builder.Writeln("Inline MathML equations:");

        // Sample MathML expressions.
        string[] mathMlExpressions = new string[]
        {
            "<math><mi>a</mi><mo>+</mo><mi>b</mi><mo>=</mo><mi>c</mi></math>",
            "<math><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup><mo>=</mo><msup><mi>z</mi><mn>2</mn></msup></math>"
        };

        // Process each MathML expression.
        foreach (string mathMl in mathMlExpressions)
        {
            // Convert MathML to a readable string.
            string readableEquation = ConvertMathMlToReadable(mathMl);

            // Generate an SVG that displays the readable equation.
            string svgContent = GenerateSvg(readableEquation, 120, 30);
            string svgPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");
            File.WriteAllText(svgPath, svgContent);

            // Insert the SVG as an inline image shape.
            Shape shape = builder.InsertImage(svgPath);
            shape.WrapType = WrapType.Inline;          // Ensure the shape is inline.
            shape.Width = 120;                         // Normalise width.
            shape.Height = 30;                         // Normalise height.

            // Add a space after the shape for visual separation.
            builder.Write(" ");

            // Clean up the temporary SVG file.
            File.Delete(svgPath);
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MathMLShapes.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }

    // Very basic conversion: strip tags and keep inner text.
    private static string ConvertMathMlToReadable(string mathMl)
    {
        var withoutTags = Regex.Replace(mathMl, "<[^>]+>", string.Empty);
        return withoutTags.Trim();
    }

    // Generates a minimal SVG containing the provided text.
    private static string GenerateSvg(string text, int width, int height)
    {
        return $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>
<svg xmlns=""http://www.w3.org/2000/svg"" width=""{width}"" height=""{height}"">
  <rect width=""100%"" height=""100%"" fill=""white"" stroke=""black"" stroke-width=""0.5""/>
  <text x=""50%"" y=""50%"" dominant-baseline=""middle"" text-anchor=""middle"" font-family=""Arial"" font-size=""12"">{System.Security.SecurityElement.Escape(text)}</text>
</svg>";
    }
}
