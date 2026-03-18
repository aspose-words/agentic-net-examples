using System;
using System.IO;
using System.Text;
using Aspose.Words;

namespace AsposeWordsPlainTextExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Word document.
            string inputFilePath = @"C:\Docs\SourceDocument.docx";

            // Path where the plain‑text output will be saved.
            string outputFilePath = @"C:\Docs\PlainTextOutput.txt";

            // Load the document from the file system.
            Document doc = new Document(inputFilePath);

            // StringBuilder to accumulate the text of each section.
            StringBuilder plainTextBuilder = new StringBuilder();

            // Iterate through all sections in the document.
            foreach (Section section in doc.Sections)
            {
                // Extract the text covered by the section's range.
                // Trim to remove leading/trailing control characters.
                string sectionText = section.Range.Text.Trim();

                // Append the section text followed by a line break.
                plainTextBuilder.AppendLine(sectionText);
            }

            // Write the concatenated plain‑text to the output file.
            File.WriteAllText(outputFilePath, plainTextBuilder.ToString());

            // Optional: inform the user that the operation completed.
            Console.WriteLine("Plain‑text extraction completed. Output saved to:");
            Console.WriteLine(outputFilePath);
        }
    }
}
