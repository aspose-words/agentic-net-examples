using System;
using System.IO;
using System.Text;
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
        string[] mathMlExpressions = new[]
        {
            "<math><mi>a</mi><mo>+</mo><mi>b</mi><mo>=</mo><mi>c</mi></math>",
            "<math><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup><mo>=</mo><msup><mi>z</mi><mn>2</mn></msup></math>"
        };

        // Insert each expression as an inline SVG image.
        foreach (string mathMl in mathMlExpressions)
        {
            // Very simple conversion: strip tags and keep the inner text.
            // In a real scenario you would parse MathML properly.
            string equation = StripMathMlTags(mathMl);

            // Build a minimal SVG that displays the equation text.
            string svgContent = $@"<svg xmlns=""http://www.w3.org/2000/svg"" width=""120"" height=""30"">
  <text x=""0"" y=""20"" font-family=""Arial"" font-size=""14"">{System.Security.SecurityElement.Escape(equation)}</text>
</svg>";

            // Convert SVG string to a UTF‑8 stream.
            using (MemoryStream svgStream = new MemoryStream(Encoding.UTF8.GetBytes(svgContent)))
            {
                // Insert the SVG as an inline image shape.
                Shape shape = builder.InsertImage(svgStream);
                shape.WrapType = WrapType.Inline; // Ensure inline positioning.
                shape.Width = 120;                 // Normalized width.
                shape.Height = 30;                 // Normalized height.
            }

            // Add a space between equations.
            builder.Write(" ");
        }

        // Adjust baseline alignment for the paragraph containing the shapes.
        builder.CurrentParagraph.ParagraphFormat.BaselineAlignment = BaselineAlignment.Baseline;

        // Save the document.
        string outputPath = "MathMLShapes.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file '{outputPath}'.");
    }

    // Helper that removes simple MathML tags, leaving only the inner characters.
    private static string StripMathMlTags(string mathMl)
    {
        var sb = new StringBuilder();
        bool insideTag = false;
        foreach (char c in mathMl)
        {
            if (c == '<')
                insideTag = true;
            else if (c == '>')
                insideTag = false;
            else if (!insideTag)
                sb.Append(c);
        }
        return sb.ToString().Trim();
    }
}
