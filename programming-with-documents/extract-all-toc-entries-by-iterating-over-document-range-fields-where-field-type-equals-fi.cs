using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field.
        // The switches specify that headings level 1‑3 will be included.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings that will become TOC entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Update fields so the TOC is populated.
        doc.UpdateFields();

        // Extract all TOC entries.
        List<string> tocEntries = new List<string>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
            {
                // Cast to FieldToc to access the displayed result.
                FieldToc tocField = (FieldToc)field;
                // The DisplayResult contains the TOC text (lines separated by line breaks).
                string result = tocField.DisplayResult;
                // Split into individual lines and add to the list.
                string[] lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                tocEntries.AddRange(lines);
            }
        }

        // Output the extracted TOC entries.
        Console.WriteLine("Extracted TOC entries:");
        foreach (string entry in tocEntries)
        {
            Console.WriteLine(entry);
        }

        // Save the document (optional, demonstrates saving).
        doc.Save("ExtractToc.docx");
    }
}
