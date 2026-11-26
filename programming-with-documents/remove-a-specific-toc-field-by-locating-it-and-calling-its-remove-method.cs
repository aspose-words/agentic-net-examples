using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace RemoveTocExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents field.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Add some headings that would appear in the TOC.
            builder.Font.Size = 16;
            builder.Font.Bold = true;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.Font.Size = 14;
            builder.Font.Bold = false;
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

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
            const string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
