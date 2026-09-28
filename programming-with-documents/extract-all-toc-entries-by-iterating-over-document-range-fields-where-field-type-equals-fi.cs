using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace TocExtractor
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents field with typical switches.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.Writeln(); // Add a blank line after the TOC.

            // Add headings that will be captured by the TOC.

            // Heading 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.Writeln("Some content for chapter 1.");

            // Heading 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("Details of section 1.1.");

            // Another Heading 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.Writeln("Some content for chapter 2.");

            // Update fields so the TOC is generated.
            doc.UpdateFields();

            // Save the document (optional, just to demonstrate creation).
            const string outputPath = "Sample.docx";
            doc.Save(outputPath);

            // Iterate over all fields and extract TOC entries.
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // The result of the TOC field contains the generated entries.
                    string tocResult = field.Result;
                    Console.WriteLine("TOC Entries:");
                    Console.WriteLine(tocResult);
                }
            }
        }
    }
}
