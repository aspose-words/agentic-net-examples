using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define the output file path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineHeight.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150 % of the default (default is 12 pt, so 1.5 × 12 pt = 18 pt).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 18 points = 150 % line height.

        // Insert a paragraph with some sample text.
        builder.Writeln("This paragraph has a custom line height of 150 %.");

        // Save the document.
        doc.Save(outputPath);

        // Reload the document to verify that the formatting was saved correctly.
        Document loadedDoc = new Document(outputPath);
        Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;

        // Verify the line spacing rule and value.
        bool isCorrectRule = firstParagraph.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        bool isCorrectSpacing = Math.Abs(firstParagraph.ParagraphFormat.LineSpacing - 18) < 0.01;

        // Output verification result.
        Console.WriteLine(isCorrectRule && isCorrectSpacing
            ? "Line height set correctly to 150 %."
            : "Line height verification failed.");
    }
}
