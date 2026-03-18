using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with default line spacing.
        builder.Writeln("This paragraph uses the default line spacing.");

        // Add a new paragraph that we will modify.
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the line spacing rule to Multiple (interpreted as a multiple of the default line height)
        // and specify the multiple as 1.5.
        builder.CurrentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Save the document to a file.
        doc.Save("ParagraphLineSpacing.docx");
    }
}
