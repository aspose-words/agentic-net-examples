using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and add a couple of paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with custom line spacing.");
        builder.Writeln("Second paragraph with custom line spacing.");

        // Apply a non‑default line spacing to demonstrate the reset later.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            para.ParagraphFormat.LineSpacing = 30; // 30 points exact spacing.
        }

        // Reset line spacing for all paragraphs to the default.
        // Setting LineSpacing to 0 tells Word to use its default line spacing.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.LineSpacing = 0;
            // Clear any explicit rule so the default (Multiple) is applied.
            para.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        }

        // Save the modified document.
        doc.Save("ResetLineSpacing.docx");
    }
}
