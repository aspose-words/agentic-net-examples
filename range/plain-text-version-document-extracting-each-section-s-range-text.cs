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
            // Use a temporary folder for the demo files.
            string tempFolder = Path.GetTempPath();
            string inputFilePath = Path.Combine(tempFolder, "SourceDocument.docx");
            string outputFilePath = Path.Combine(tempFolder, "PlainTextOutput.txt");

            // If the source document does not exist, create a simple one.
            if (!File.Exists(inputFilePath))
            {
                Document docToCreate = new Document();
                DocumentBuilder builder = new DocumentBuilder(docToCreate);
                builder.Writeln("First section text.");
                builder.InsertBreak(BreakType.SectionBreakNewPage);
                builder.Writeln("Second section text.");
                docToCreate.Save(inputFilePath);
            }

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

            // Inform the user that the operation completed.
            Console.WriteLine("Plain‑text extraction completed. Output saved to:");
            Console.WriteLine(outputFilePath);
        }
    }
}
