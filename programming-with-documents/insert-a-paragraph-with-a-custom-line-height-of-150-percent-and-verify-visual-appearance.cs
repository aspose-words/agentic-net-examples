using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150 % of the default (12 pt) line height.
        // Using the Multiple rule with a value of 18 pt gives 1.5 × 12 pt = 18 pt.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 150 % line height.

        // Insert a paragraph that uses the custom line height.
        builder.Writeln("This paragraph has a line height of 150 %.");

        // Save the document to disk.
        const string outputPath = "CustomLineHeight.docx";
        doc.Save(outputPath);

        // Reload the document to verify that the formatting was saved correctly.
        Document loadedDoc = new Document(outputPath);
        Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;

        // Output the applied line spacing settings.
        Console.WriteLine($"LineSpacingRule: {firstParagraph.ParagraphFormat.LineSpacingRule}");
        Console.WriteLine($"LineSpacing (points): {firstParagraph.ParagraphFormat.LineSpacing}");
    }
}
