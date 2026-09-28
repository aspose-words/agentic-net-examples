using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample MathML expressions.
        string[] mathmlExpressions = new[]
        {
            "<math><mi>a</mi><mo>+</mo><mi>b</mi></math>",
            "<math><msup><mi>x</mi><mn>2</mn></msup><mo>=</mo><mn>4</mn></math>",
            "<math><mfrac><mi>1</mi><mi>2</mi></mfrac></math>"
        };

        // Insert each MathML expression as an inline SVG image.
        foreach (string mathml in mathmlExpressions)
        {
            // Convert MathML to a simple readable string (strip tags).
            string displayText = Regex.Replace(mathml, "<.*?>", string.Empty);

            // Create an SVG that displays the readable equation string.
            string svgContent = $@"<svg xmlns='http://www.w3.org/2000/svg' width='200' height='30'>
  <text x='0' y='20' font-family='Arial' font-size='14'>{System.Security.SecurityElement.Escape(displayText)}</text>
</svg>";

            // Save SVG to a temporary file.
            string tempSvgPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");
            File.WriteAllText(tempSvgPath, svgContent);

            // Insert the SVG as an inline image shape.
            Shape shape = builder.InsertImage(tempSvgPath);

            // Ensure the shape is inline.
            shape.WrapType = WrapType.Inline;

            // Set a uniform size for consistency.
            shape.Width = 200;
            shape.Height = 30;

            // Add a space after the shape to separate consecutive images.
            builder.Write(" ");

            // Clean up the temporary SVG file.
            File.Delete(tempSvgPath);
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MathMLShapes.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
