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

        // Add a paragraph with sample text.
        builder.Writeln("This is a sample paragraph. It will have line spacing set to 1.5 lines.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Adjust line spacing to 1.5 lines.
        paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        paragraph.ParagraphFormat.LineSpacing = 1.5;

        // Validate that the line spacing was set correctly.
        if (paragraph.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(paragraph.ParagraphFormat.LineSpacing - 1.5) > 0.001)
        {
            throw new InvalidOperationException("Line spacing was not set correctly.");
        }

        // Save the document to disk.
        string outputPath = "LineSpacingExample.docx";
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
