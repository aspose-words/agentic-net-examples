using System;
using Aspose.Words;

class DoubleLineSpacingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which simplifies adding content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the line spacing rule to "Multiple" so that the LineSpacing value
        // is interpreted as a multiple of the default line height.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

        // Assign a value of 2.0 to achieve double line spacing.
        builder.ParagraphFormat.LineSpacing = 2.0;

        // Add a paragraph to demonstrate the double line spacing.
        builder.Writeln("This paragraph is formatted with double line spacing.");

        // Save the document to disk.
        doc.Save("DoubleLineSpacing.docx");
    }
}
