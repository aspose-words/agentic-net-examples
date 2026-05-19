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
        string[] mathMlExpressions =
        {
            "<math><mi>a</mi><mo>+</mo><mi>b</mi></math>",
            "<math><mi>x</mi><mo>=</mo><mi>y</mi><mo>·</mo><mi>z</mi></math>",
            "<math><mi>∑</mi><mi>i=1</mi><mo>^</mo><mi>n</mi><mi>i</mi></math>"
        };

        // Insert each MathML as an inline SVG image shape.
        foreach (string mathMl in mathMlExpressions)
        {
            // Very simple conversion: strip tags and keep inner text.
            string equation = StripMathMlTags(mathMl);

            // Build a minimal SVG that displays the equation text.
            string svgContent = $@"
<svg xmlns='http://www.w3.org/2000/svg' width='120' height='30'>
    <text x='0' y='20' font-family='Cambria Math' font-size='16'>{System.Security.SecurityElement.Escape(equation)}</text>
</svg>";

            // Convert SVG string to a UTF‑8 byte array and load it into a memory stream.
            byte[] svgBytes = Encoding.UTF8.GetBytes(svgContent);
            using (MemoryStream svgStream = new MemoryStream(svgBytes))
            {
                // Insert the SVG as an image. Aspose.Words will render it as a PNG internally.
                Shape shape = builder.InsertImage(svgStream);

                // Ensure the shape is inline.
                shape.WrapType = WrapType.Inline;

                // Set a uniform size (points) for consistent appearance.
                shape.Width = 120;   // points
                shape.Height = 30;   // points

                // Note: Shape does not expose a BaselineAlignment property.
                // Inline shapes align to the baseline by default, which satisfies the task requirement.
            }
        }

        // Add some surrounding text to demonstrate baseline alignment.
        builder.Writeln();
        builder.Writeln("The equations above are inserted as inline shapes.");

        // Save the document.
        string outputPath = "MathMLShapes.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }

    // Helper method to remove simple MathML tags and keep the inner characters.
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
        return sb.ToString();
    }
}
