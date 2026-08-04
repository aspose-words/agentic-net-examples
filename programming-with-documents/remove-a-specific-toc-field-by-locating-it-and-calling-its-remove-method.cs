using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Add some headings that will appear in the TOC.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("Section 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            // Update fields so the TOC is built (optional for this example).
            doc.UpdateFields();

            // Save the document with the TOC for reference.
            doc.Save("DocumentWithToc.docx");

            // Locate the first TOC field in the document.
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
