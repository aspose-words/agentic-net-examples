using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph and add it to the document body.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Add a run with sample text.
        Run run = new Run(doc, "This is a sample paragraph with 1.5 line spacing.");
        paragraph.AppendChild(run);

        // Adjust line spacing to 1.5 lines by setting ParagraphFormat.
        paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        paragraph.ParagraphFormat.LineSpacing = 1.5;

        // Validate that the line spacing was set correctly.
        if (Math.Abs(paragraph.ParagraphFormat.LineSpacing - 1.5) > 0.0001 ||
            paragraph.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple)
        {
            throw new InvalidOperationException("Failed to set line spacing to 1.5.");
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
