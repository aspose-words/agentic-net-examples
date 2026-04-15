using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsRemoveToc
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field.
            // The switches configure the TOC to include heading levels 1‑3 and add hyperlinks.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Add some headings so the TOC has entries.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("Section 1.2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 2.1");

            // Update fields so the TOC displays its entries.
            doc.UpdateFields();

            // Save the document before removal (optional, just for reference).
            doc.Save("DocumentWithToc.docx");

            // Locate the TOC field in the document.
            Field tocField = null;
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    tocField = field;
                    break;
                }
            }

            // If a TOC field was found, remove it.
            if (tocField != null)
            {
                tocField.Remove();
            }

            // Save the document after the TOC has been removed.
            doc.Save("DocumentWithoutToc.docx");
        }
    }
}
