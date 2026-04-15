using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field at the beginning of the document.
        builder.InsertField(FieldType.FieldTOC, true);
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings that the TOC will pick up.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.InsertBreak(BreakType.PageBreak);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.InsertBreak(BreakType.PageBreak);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.InsertBreak(BreakType.PageBreak);

        // Update all fields so the TOC is populated.
        doc.UpdateFields();

        // Iterate over all fields in the document and extract TOC entries.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // The Result property contains the visible TOC text.
                Console.WriteLine("=== TOC Entries ===");
                Console.WriteLine(field.Result);
            }
        }

        // Save the document to the current directory (optional).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath);
    }
}
