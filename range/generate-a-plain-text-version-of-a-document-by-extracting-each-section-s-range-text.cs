using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace RangePlainTextExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with two sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First section content.
            builder.Writeln("Section 1 - Paragraph 1");
            builder.Writeln("Section 1 - Paragraph 2");

            // Insert a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second section content.
            builder.Writeln("Section 2 - Paragraph 1");
            builder.Writeln("Section 2 - Paragraph 2");

            // Save the sample document locally.
            string docPath = "Sample.docx";
            doc.Save(docPath);

            // Load the document from the saved file.
            Document loadedDoc = new Document(docPath);

            // Extract plain text from each section's Range.Text.
            StringBuilder plainTextBuilder = new StringBuilder();

            foreach (Section section in loadedDoc.Sections)
            {
                // Trim to remove leading/trailing whitespace that may be added by Aspose.Words.
                string sectionText = section.Range.Text.Trim();
                plainTextBuilder.AppendLine(sectionText);
                // Optional separator between sections.
                plainTextBuilder.AppendLine("--- End of Section ---");
                plainTextBuilder.AppendLine();
            }

            // Write the extracted plain text to a .txt file.
            string txtPath = "PlainText.txt";
            File.WriteAllText(txtPath, plainTextBuilder.ToString());

            // Indicate completion.
            Console.WriteLine($"Plain‑text extraction completed. Output saved to '{txtPath}'.");
        }
    }
}
