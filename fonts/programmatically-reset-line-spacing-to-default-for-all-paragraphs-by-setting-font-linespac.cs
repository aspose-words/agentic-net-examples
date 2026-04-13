using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a custom line spacing (multiple of default).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 24; // 2 lines (default is 12 points).
        builder.Writeln("Paragraph with custom line spacing.");

        // Add another paragraph with a different line spacing rule.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
        builder.ParagraphFormat.LineSpacing = 30;
        builder.Writeln("Another paragraph with custom line spacing.");

        // Save the document before resetting (optional, just to have an intermediate file).
        string initialPath = "InitialLineSpacing.docx";
        doc.Save(initialPath);

        // Reset line spacing for all paragraphs to the default value.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Setting LineSpacing to 0 and using the Multiple rule restores the default spacing.
            paragraph.ParagraphFormat.LineSpacing = 0;
            paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        }

        // Save the document after resetting line spacing.
        string resetPath = "ResetLineSpacing.docx";
        doc.Save(resetPath);

        // Verify that the output file exists.
        if (!File.Exists(resetPath))
        {
            throw new Exception("The document was not saved correctly.");
        }
    }
}
