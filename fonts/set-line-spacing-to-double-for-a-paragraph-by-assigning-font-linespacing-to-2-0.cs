using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph with double line spacing.");

        // Retrieve the first paragraph in the document.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Set line spacing to double (2.0) using ParagraphFormat.
        paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        paragraph.ParagraphFormat.LineSpacing = 2.0;

        // Validate that the line spacing was set correctly.
        if (paragraph.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(paragraph.ParagraphFormat.LineSpacing - 2.0) > 0.0001)
        {
            throw new InvalidOperationException("Line spacing was not set to the expected value.");
        }

        // Save the document to disk.
        string outputPath = "OutputDoubleLineSpacing.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
