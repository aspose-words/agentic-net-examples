using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with default line spacing.
        builder.Writeln("This paragraph uses the default line spacing.");

        // Set line spacing to double for subsequent paragraphs.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 2.0;

        // Add a paragraph that will inherit the double line spacing.
        builder.Writeln("This paragraph has double line spacing.");

        // Validate that the line spacing was set correctly.
        if (builder.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(builder.ParagraphFormat.LineSpacing - 2.0) > 0.0001)
        {
            throw new InvalidOperationException("Line spacing was not set to double.");
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
