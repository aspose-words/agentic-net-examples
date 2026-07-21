using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the paragraph's line spacing to 1.5 lines.
        // Use the Multiple rule where the value represents a multiple of the default line height (12 points).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 1.5; // 1.5 * 12 points = 18 points.

        // Validate that the line spacing rule and value were applied.
        if (builder.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(builder.ParagraphFormat.LineSpacing - 1.5) > 0.0001)
        {
            throw new InvalidOperationException("Failed to set paragraph line spacing to 1.5 lines.");
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }

        // Optionally, inform the user (no interactive input required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
