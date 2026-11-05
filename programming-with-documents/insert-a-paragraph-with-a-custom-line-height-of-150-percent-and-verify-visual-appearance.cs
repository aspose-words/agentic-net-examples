using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150 % (multiple of 1.5).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 1.5; // 150%

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph has a line height of 150 %.");

        // Save the document.
        string outputPath = "ParagraphLineHeight.docx";
        doc.Save(outputPath);

        // Reload the document and verify the line spacing.
        Document loadedDoc = new Document(outputPath);
        Paragraph para = loadedDoc.FirstSection.Body.FirstParagraph;

        bool ruleCorrect = para.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        double actualLineSpacing = para.ParagraphFormat.LineSpacing;
        double expectedLineSpacing = 1.5;
        bool spacingCorrect = Math.Abs(actualLineSpacing - expectedLineSpacing) < 0.001;

        Console.WriteLine(ruleCorrect && spacingCorrect
            ? "Line spacing set correctly."
            : $"Line spacing mismatch. Expected rule Multiple and spacing {expectedLineSpacing}, got rule {para.ParagraphFormat.LineSpacingRule} and spacing {actualLineSpacing}.");
    }
}
