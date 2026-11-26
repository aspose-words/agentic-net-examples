using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceHeaderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a header containing text to replace.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert header text.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Company: OldName");
            builder.Writeln();
            builder.Write("Report generated on DATE");

            // Add a body paragraph.
            builder.MoveToDocumentEnd();
            builder.Writeln("This is the body of the document.");

            // Save the sample document.
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document from the saved file.
            Document loadedDoc = new Document(inputPath);

            // Replace text within all primary headers.
            int totalReplacements = 0;
            foreach (Section section in loadedDoc.Sections)
            {
                HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                {
                    int replaced = header.Range.Replace("OldName", "NewName", new FindReplaceOptions());
                    totalReplacements += replaced;
                }
            }

            if (totalReplacements == 0)
                throw new InvalidOperationException("No replacements were made in the headers.");

            // Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
