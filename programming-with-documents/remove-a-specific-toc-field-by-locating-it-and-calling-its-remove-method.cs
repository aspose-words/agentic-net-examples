using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings so the TOC has entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 1.1");
        builder.Writeln("Heading 1.2");

        // Update fields to populate the TOC.
        doc.UpdateFields();

        // Locate the first TOC field (FieldToc) in the document.
        FieldToc tocField = null;
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                tocField = (FieldToc)field;
                break;
            }
        }

        // Remove the TOC field if it was found.
        if (tocField != null)
        {
            tocField.Remove();
        }

        // Save the document after removal.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }
}
