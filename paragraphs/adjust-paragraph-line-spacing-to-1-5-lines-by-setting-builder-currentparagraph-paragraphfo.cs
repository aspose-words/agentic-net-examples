using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This is a sample paragraph with 1.5 line spacing.");

        // Adjust the line spacing of the current paragraph.
        // Set the rule to Multiple so that the value is interpreted as a multiple of the default line height.
        builder.CurrentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        // According to the task, set LineSpacing to 1.5 (interpreted as 1.5 lines because of the rule above).
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Save the document to the local file system.
        const string outputFile = "ParagraphLineSpacing.docx";
        doc.Save(outputFile);
    }
}
