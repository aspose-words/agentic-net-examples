using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and set formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the line spacing rule to Multiple so the value is interpreted as a number of lines.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

        // Assign 1.5 lines as the line spacing.
        builder.ParagraphFormat.LineSpacing = 1.5;

        // Add a paragraph to demonstrate the spacing.
        builder.Writeln("This paragraph uses 1.5 line spacing.");

        // Save the document.
        doc.Save("ParagraphLineSpacing.docx");
    }
}
