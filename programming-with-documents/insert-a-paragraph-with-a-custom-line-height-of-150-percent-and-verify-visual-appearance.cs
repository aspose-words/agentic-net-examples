using System;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150% (multiple line spacing rule).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 1.5;

        // Insert a paragraph with the custom line height.
        builder.Writeln("This paragraph has a line spacing of 150 percent.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Reload the document to verify the line spacing.
        Document loadedDoc = new Document(outputPath);
        Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;

        bool isCorrectRule = firstParagraph.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        bool isCorrectValue = Math.Abs(firstParagraph.ParagraphFormat.LineSpacing - 1.5) < 0.0001;

        if (isCorrectRule && isCorrectValue)
        {
            Console.WriteLine("Verification passed: line spacing is 150% as expected.");
        }
        else
        {
            Console.WriteLine("Verification failed: line spacing does not match expected value.");
        }
    }
}
