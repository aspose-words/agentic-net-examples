using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and format the paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the line spacing rule to Multiple (i.e., a multiple of the default line height).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

        // Set the line spacing to double the default (default line height is 12 points, so 24 points = double spacing).
        builder.ParagraphFormat.LineSpacing = 24.0;

        // Write a sample line to demonstrate the double line spacing.
        builder.Writeln("This paragraph has double line spacing.");

        // Validate that the settings were applied correctly.
        if (builder.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(builder.ParagraphFormat.LineSpacing - 24.0) > 0.001)
        {
            throw new InvalidOperationException("Failed to set double line spacing on the paragraph.");
        }

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DoubleLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }

        // Indicate success (no interactive prompts).
        Console.WriteLine("Document saved successfully to: " + outputPath);
    }
}
