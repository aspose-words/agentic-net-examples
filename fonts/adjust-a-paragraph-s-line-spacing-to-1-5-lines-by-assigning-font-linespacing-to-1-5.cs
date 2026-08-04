using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the line spacing rule to Multiple (number of lines) and specify 1.5 lines.
        // One line equals 12 points, so 1.5 lines = 18 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18.0;

        // Validate that the line spacing was applied.
        double appliedSpacing = builder.ParagraphFormat.LineSpacing;
        if (Math.Abs(appliedSpacing - 18.0) > 0.001)
        {
            throw new InvalidOperationException("Line spacing was not set correctly.");
        }

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
