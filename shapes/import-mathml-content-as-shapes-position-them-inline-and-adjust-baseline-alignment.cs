using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Sample MathML expressions.
        string[] mathmlExpressions = new[]
        {
            "<math><mi>x</mi><mo>=</mo><mfrac><mi>-b</mi><msqrt><msup><mi>b</mi><mn>2</mn></msup><mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></msqrt></mfrac></math>",
            "<math><msup><mi>e</mi><mi>iπ</mi></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>"
        };

        // Folder for temporary SVG files.
        string tempDir = Path.Combine(Path.GetTempPath(), "MathMLSvgDemo");
        Directory.CreateDirectory(tempDir);

        // Create simple SVG files that display a readable version of each MathML expression.
        string[] svgPaths = new string[mathmlExpressions.Length];
        for (int i = 0; i < mathmlExpressions.Length; i++)
        {
            string svgPath = Path.Combine(tempDir, $"eq{i + 1}.svg");
            // Convert the MathML to a plain‑text representation for the demo.
            string displayText = System.Security.SecurityElement.Escape(
                System.Text.RegularExpressions.Regex.Replace(
                    mathmlExpressions[i],
                    "<[^>]+>",
                    string.Empty));

            string svgContent = $@"<?xml version=""1.0"" encoding=""UTF-8""?>
<svg xmlns=""http://www.w3.org/2000/svg"" width=""300"" height=""30"">
  <text x=""0"" y=""20"" font-family=""Arial"" font-size=""14"">{displayText}</text>
</svg>";
            File.WriteAllText(svgPath, svgContent);
            svgPaths[i] = svgPath;
        }

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("Below are inline MathML equations rendered as SVG images:");

        // Paragraph that will contain the inline SVG images.
        builder.InsertParagraph();

        // Insert each SVG as an inline image shape.
        for (int i = 0; i < svgPaths.Length; i++)
        {
            if (i > 0)
                builder.Write(" "); // Space between images.

            Shape shape = builder.InsertImage(svgPaths[i]);
            shape.Width = 300;
            shape.Height = 30;
            shape.WrapType = WrapType.Inline; // Ensure the shape behaves as an inline object.
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MathMLShapes.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Clean up temporary SVG files.
        foreach (var path in svgPaths)
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        if (Directory.Exists(tempDir))
            Directory.Delete(tempDir, true);
    }
}
