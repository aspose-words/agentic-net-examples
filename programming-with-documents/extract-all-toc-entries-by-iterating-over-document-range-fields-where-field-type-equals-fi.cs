using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace ExtractTocEntries
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field.
            // The switches configure the TOC to include heading levels 1‑3 and to create hyperlinks.
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
            builder.Writeln("Subsection 2.1.2");

            // Update all fields so that the TOC is populated.
            doc.UpdateFields();

            // Save the document – this satisfies the requirement to produce an output file.
            string docPath = "TocDocument.docx";
            doc.Save(docPath);

            // Iterate over all fields in the document and extract entries from TOC fields.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // Cast to FieldToc to access TOC‑specific properties.
                    FieldToc tocField = (FieldToc)field;

                    // The Result property contains the displayed TOC text (entries).
                    string tocResult = tocField.Result;

                    Console.WriteLine("=== Extracted TOC Entries ===");
                    // The result may contain line breaks; split for clearer output.
                    string[] lines = tocResult.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        Console.WriteLine(line.Trim());
                    }
                }
            }
        }
    }
}
