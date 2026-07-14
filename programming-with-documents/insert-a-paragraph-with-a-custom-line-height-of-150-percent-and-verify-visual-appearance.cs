using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150% (1.5 lines). One line = 12 points, so 1.5 lines = 18 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 12 * 1.5 = 18 points

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph has a custom line height of 150%.");

        // Save the document to the local file system.
        string outputPath = "ParagraphLineHeight.docx";
        doc.Save(outputPath);

        // Verify that the paragraph's formatting was applied correctly.
        Paragraph para = doc.FirstSection.Body.Paragraphs[0];
        bool ruleCorrect = para.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        bool spacingCorrect = Math.Abs(para.ParagraphFormat.LineSpacing - 18) < 0.001;

        Console.WriteLine($"Line spacing rule set correctly: {ruleCorrect}");
        Console.WriteLine($"Line spacing value set correctly: {spacingCorrect}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
