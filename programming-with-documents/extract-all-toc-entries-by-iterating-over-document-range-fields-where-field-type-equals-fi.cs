using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace TocExtractor
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents field.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Add some headings that will appear in the TOC.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 2.1");

            // Update fields so the TOC is populated.
            doc.UpdateFields();

            // Save the document (optional, just to have an output file).
            doc.Save("Sample.docx");

            // Iterate over all fields and extract TOC entries.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // The Result property contains the rendered TOC text.
                    string tocText = field.Result;

                    // Output each line of the TOC.
                    string[] lines = tocText.Split(new[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        Console.WriteLine(line.Trim());
                    }
                }
            }
        }
    }
}
