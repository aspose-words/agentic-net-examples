using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Newtonsoft.Json; // Included as per required packages.

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph's line spacing to double.
        // The default line height is 12 points, so double spacing is 24 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 24;

        // Add some text to demonstrate the line spacing.
        builder.Writeln("This paragraph has double line spacing.");
        builder.Writeln("Second line of the same paragraph.");

        // Validate that the line spacing was applied correctly.
        if (builder.ParagraphFormat.LineSpacing != 24 ||
            builder.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple)
        {
            throw new InvalidOperationException("Line spacing was not set correctly.");
        }

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DoubleLineSpacing.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved.", outputPath);
        }
    }
}
