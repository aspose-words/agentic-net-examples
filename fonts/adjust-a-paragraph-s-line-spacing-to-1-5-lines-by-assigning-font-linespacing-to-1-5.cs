using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a sample paragraph.
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the line spacing rule to Multiple (number of lines).
        builder.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Multiple;

        // One line equals 12 points; 1.5 lines = 18 points.
        builder.ParagraphFormat.LineSpacing = 18.0;

        // Validate that the line spacing was set correctly.
        if (builder.ParagraphFormat.LineSpacingRule != Aspose.Words.LineSpacingRule.Multiple ||
            Math.Abs(builder.ParagraphFormat.LineSpacing - 18.0) > 0.001)
        {
            throw new InvalidOperationException("Line spacing was not set to 1.5 lines.");
        }

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
