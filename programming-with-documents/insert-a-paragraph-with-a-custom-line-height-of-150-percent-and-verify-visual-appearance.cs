using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150% (multiple line spacing)
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 1.5; // 150%

        // Insert a paragraph with sample text
        builder.Writeln("This paragraph has a line height of 150 percent.");

        // Save the document
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Reload the document to verify the line spacing
        Document loadedDoc = new Document(outputPath);
        Paragraph paragraph = loadedDoc.FirstSection.Body.FirstParagraph;

        bool ruleIsMultiple = paragraph.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        bool spacingIs150 = Math.Abs(paragraph.ParagraphFormat.LineSpacing - 1.5) < 0.001;

        if (ruleIsMultiple && spacingIs150)
        {
            Console.WriteLine("Verification passed: Paragraph line spacing is set to 150%.");
        }
        else
        {
            Console.WriteLine("Verification failed: Paragraph line spacing is not as expected.");
        }
    }
}
