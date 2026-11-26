using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Sample MathML expressions.
        var mathMlExpressions = new List<string>
        {
            "<math><mi>x</mi><mo>+</mo><mi>y</mi></math>",
            "<math><msup><mi>a</mi><mn>2</mn></msup></math>",
            "<math><mfrac><mi>1</mi><mi>2</mi></mfrac></math>"
        };

        // Create a new document and builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Start a new paragraph for inline shapes.
        builder.Writeln();
        builder.MoveToDocumentEnd();

        // Temporary folder for generated SVG files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "MathMLSvgTemp");
        Directory.CreateDirectory(tempFolder);

        // Insert each MathML as an inline SVG image shape.
        for (int i = 0; i < mathMlExpressions.Count; i++)
        {
            string mathMl = mathMlExpressions[i];

            // Create a simple SVG that displays a readable representation of the MathML.
            string svgContent = $@"<svg xmlns='http://www.w3.org/2000/svg' width='200' height='30'>
  <text x='0' y='20' font-family='Arial' font-size='14'>{System.Security.SecurityElement.Escape(mathMl)}</text>
</svg>";

            string svgPath = Path.Combine(tempFolder, $"math{i}.svg");
            File.WriteAllText(svgPath, svgContent);

            // Insert the SVG as an inline image shape.
            Shape shape = builder.InsertImage(svgPath);
            shape.Width = 200;
            shape.Height = 30;
            shape.WrapType = WrapType.Inline;

            // Add a space after each shape for readability.
            builder.Write(" ");
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MathMLShapes.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Clean up temporary SVG files.
        try
        {
            Directory.Delete(tempFolder, true);
        }
        catch
        {
            // Ignored – temporary files are not critical.
        }
    }
}
