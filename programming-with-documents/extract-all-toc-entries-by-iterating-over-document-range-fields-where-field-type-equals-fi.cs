using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsTocExtraction
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

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 2.1");

            // Update all fields so the TOC is populated.
            doc.UpdateFields();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string docPath = Path.Combine(outputDir, "SampleToc.docx");
            doc.Save(docPath);

            // Iterate over all fields and extract TOC entries.
            Console.WriteLine("Extracted TOC entries:");
            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldTOC)
                {
                    // Cast to FieldToc to access TOC‑specific properties.
                    FieldToc tocField = (FieldToc)field;

                    // The DisplayResult property contains the rendered TOC text.
                    string tocResult = tocField.DisplayResult?.Trim();

                    if (!string.IsNullOrEmpty(tocResult))
                    {
                        // Split the result into individual lines (entries).
                        string[] entries = tocResult.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string entry in entries)
                        {
                            Console.WriteLine(entry);
                        }
                    }
                }
            }
        }
    }
}
