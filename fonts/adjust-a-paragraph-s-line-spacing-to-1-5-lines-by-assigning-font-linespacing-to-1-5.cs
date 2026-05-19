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

        // Write a paragraph.
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the line spacing rule to Multiple (number of lines).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

        // Default line spacing is 12 points. 1.5 lines = 12 * 1.5 = 18 points.
        builder.ParagraphFormat.LineSpacing = 18;

        // Validate that the line spacing was set correctly.
        if (builder.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(builder.ParagraphFormat.LineSpacing - 18) > 0.001)
        {
            throw new InvalidOperationException("Line spacing was not set correctly.");
        }

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ParagraphLineSpacing.docx");
        doc.Save(outputPath);

        // Ensure the file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved.", outputPath);
        }

        // Indicate completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
