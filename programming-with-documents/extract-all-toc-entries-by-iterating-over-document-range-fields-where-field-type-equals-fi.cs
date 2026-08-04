using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents (TOC) field.
        // The switches configure the TOC to include heading levels 1‑3 and create hyperlinks.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings that will be picked up by the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1");

        // Update all fields so the TOC reflects the headings.
        doc.UpdateFields();

        // Save the document to the Artifacts folder.
        string docPath = Path.Combine(artifactsDir, "DocumentWithToc.docx");
        doc.Save(docPath);

        // Iterate over all fields in the document and extract TOC entries.
        Console.WriteLine("Extracted TOC entries:");
        foreach (Field field in doc.Range.Fields)
        {
            // Check if the field is a TOC field.
            if (field.Type == FieldType.FieldTOC)
            {
                // Cast to FieldToc to access TOC‑specific properties.
                FieldToc tocField = (FieldToc)field;

                // The DisplayResult property contains the rendered TOC text (entries).
                string tocText = tocField.DisplayResult?.Trim();

                if (!string.IsNullOrEmpty(tocText))
                {
                    // Print each line of the TOC.
                    string[] lines = tocText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        Console.WriteLine(line);
                    }
                }
            }
        }
    }
}
