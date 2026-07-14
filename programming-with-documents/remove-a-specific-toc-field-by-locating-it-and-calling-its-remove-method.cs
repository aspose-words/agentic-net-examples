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
            // The switches specify which heading levels to include and enable hyperlinks.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Add some headings so the TOC would have entries if it were updated.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("Section 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            // Update fields to generate the TOC content (optional for this example).
            doc.UpdateFields();

            // Locate the first TOC field in the document.
            FieldToc tocField = null;
            foreach (Field field in doc.Range.Fields)
            {
                if (field is FieldToc)
                {
                    tocField = (FieldToc)field;
                    break;
                }
            }

            // If a TOC field was found, remove it from the document.
            if (tocField != null)
            {
                tocField.Remove();
            }

            // Save the resulting document.
            string outputPath = "RemovedToc.docx";
            doc.Save(outputPath);
        }
    }
}
