using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph to the document.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Add a run with sample text.
        Run run = new Run(doc, "This is a sample paragraph with double line spacing.");
        paragraph.AppendChild(run);

        // Set line spacing to double (2.0) via ParagraphFormat.
        paragraph.ParagraphFormat.LineSpacing = 2.0;
        paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

        // Validate that the line spacing was set correctly.
        if (Math.Abs(paragraph.ParagraphFormat.LineSpacing - 2.0) > 0.0001 ||
            paragraph.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple)
        {
            throw new InvalidOperationException("Line spacing was not set to double.");
        }

        // Save the document.
        string outputPath = "OutputDoubleLineSpacing.docx";
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
