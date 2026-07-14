using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Adjust the line spacing of the current paragraph to 1.5 lines.
        // Set the rule to Multiple so the value is interpreted as a multiple of the default line spacing.
        builder.CurrentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Add another paragraph to demonstrate that the previous paragraph keeps its custom spacing.
        builder.Writeln("Another paragraph with default spacing.");

        // Save the document to the local file system.
        doc.Save("ParagraphLineSpacing.docx");
    }
}
