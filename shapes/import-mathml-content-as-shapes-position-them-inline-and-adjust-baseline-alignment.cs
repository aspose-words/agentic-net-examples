using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class MathMlToShapeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample MathML expressions.
        List<string> mathMlList = new List<string>
        {
            "<math><mi>a</mi><mo>+</mo><mi>b</mi></math>",
            "<math><mi>x</mi><mo>=</mo><mi>y</mi><mo>²</mo></math>",
            "<math><mi>∑</mi><mi>i=1</mi><mo>ⁿ</mo><mi>i</mi></math>"
        };

        // Insert each MathML expression as an inline SVG image shape.
        foreach (string mathMl in mathMlList)
        {
            // Convert MathML to a readable string representation.
            string equation = ConvertMathMlToString(mathMl);

            // Generate an SVG stream that contains the equation text.
            using (MemoryStream svgStream = GenerateSvgStream(equation, 120, 30))
            {
                // Insert the SVG as an inline image shape.
                Shape shape = builder.InsertImage(svgStream);
                shape.Width = 120;   // Normalized width.
                shape.Height = 30;   // Normalized height.

                // Ensure the shape is treated as an inline object.
                shape.WrapType = WrapType.Inline;

                // Add a space after the shape for readability.
                builder.Write(" ");
            }
        }

        // Save the document.
        string outputPath = "MathMLShapes.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    // Simple conversion: strip all tags and keep inner text.
    private static string ConvertMathMlToString(string mathMl)
    {
        // Remove all XML tags.
        string withoutTags = Regex.Replace(mathMl, "<[^>]+>", string.Empty);
        return withoutTags.Trim();
    }

    // Generates a minimal SVG containing the provided equation text.
    private static MemoryStream GenerateSvgStream(string equation, int width, int height)
    {
        string svgTemplate = $@"<svg xmlns=""http://www.w3.org/2000/svg"" width=""{width}"" height=""{height}"">
  <text x=""0"" y=""{height - 5}"" font-family=""Cambria Math"" font-size=""16"">{System.Security.SecurityElement.Escape(equation)}</text>
</svg>";
        byte[] svgBytes = System.Text.Encoding.UTF8.GetBytes(svgTemplate);
        MemoryStream stream = new MemoryStream(svgBytes);
        stream.Position = 0;
        return stream;
    }
}
