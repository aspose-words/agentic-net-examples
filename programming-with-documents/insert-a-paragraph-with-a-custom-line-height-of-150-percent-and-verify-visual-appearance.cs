using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output file name.
        const string outputPath = "ParagraphLineSpacing.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150 % (1.5 lines). In Aspose.Words,
        // when LineSpacingRule is Multiple, the LineSpacing value is expressed
        // in points where one line equals 12 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 12 pt × 1.5 = 18 pt

        // Insert the paragraph.
        builder.Writeln("This paragraph has a line height of 150 %.");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the paragraph has the expected formatting.
        Paragraph para = doc.FirstSection.Body.Paragraphs[0];
        bool ruleCorrect = para.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        bool spacingCorrect = Math.Abs(para.ParagraphFormat.LineSpacing - 18) < 0.001;

        Console.WriteLine($"LineSpacingRule: {para.ParagraphFormat.LineSpacingRule}");
        Console.WriteLine($"LineSpacing (points): {para.ParagraphFormat.LineSpacing}");
        Console.WriteLine($"Verification: {(ruleCorrect && spacingCorrect ? "Passed" : "Failed")}");
    }
}
