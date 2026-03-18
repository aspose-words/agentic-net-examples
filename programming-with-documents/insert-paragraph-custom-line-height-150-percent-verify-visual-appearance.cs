using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the line spacing rule to Multiple (percentage of the default line height).
        // Default line height is 12 points, so 150% = 12 * 1.5 = 18 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 150% line height.

        // Insert a paragraph that will use the custom line height.
        builder.Writeln("This paragraph has a line height of 150 percent.");

        // Verify that the paragraph's formatting matches the expected values.
        Paragraph currentParagraph = builder.CurrentParagraph;
        if (currentParagraph.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(currentParagraph.ParagraphFormat.LineSpacing - 18) > 0.01)
        {
            throw new InvalidOperationException("Line spacing was not applied correctly.");
        }

        // Save the document to disk.
        doc.Save("ParagraphLineHeight150.docx");
    }
}
