using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Add some headings so the TOC would have entries.
            builder.InsertBreak(BreakType.PageBreak);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            // Update fields to generate the TOC content (optional).
            doc.UpdateFields();

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

            // Remove the TOC field if it was found.
            if (tocField != null)
            {
                tocField.Remove();
            }

            // Save the modified document.
            doc.Save("Result.docx");
        }
    }
}
