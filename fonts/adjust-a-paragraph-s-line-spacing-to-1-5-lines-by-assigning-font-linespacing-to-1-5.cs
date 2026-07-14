using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings; // For LineSpacingRule enum

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph line spacing to 1.5 lines.
        // 1 line = 12 points, so 1.5 lines = 18 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18.0;

        // Add a sample paragraph.
        builder.Writeln("This paragraph has 1.5 line spacing.");

        // Validate that the line spacing was set correctly.
        double actualSpacing = builder.ParagraphFormat.LineSpacing;
        if (Math.Abs(actualSpacing - 18.0) > 0.001)
        {
            throw new InvalidOperationException("Line spacing was not set to the expected value.");
        }

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }

        // Optionally, inform that the process completed.
        Console.WriteLine("Document saved successfully to: " + outputPath);
    }
}
