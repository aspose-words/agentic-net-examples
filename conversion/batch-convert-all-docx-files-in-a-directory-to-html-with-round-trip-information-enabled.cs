using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsBatchConvert
{
    public class Program
    {
        public static void Main()
        {
            // Define input and output directories relative to the current working directory.
            string inputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
            string outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Html");

            // Ensure the directories exist.
            Directory.CreateDirectory(inputDirectory);
            Directory.CreateDirectory(outputDirectory);

            // Create sample DOCX files if the input folder is empty.
            string[] existingDocs = Directory.GetFiles(inputDirectory, "*.docx");
            if (existingDocs.Length == 0)
            {
                for (int i = 1; i <= 2; i++)
                {
                    Document sampleDoc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                    builder.Writeln($"Sample document {i}");
                    string samplePath = Path.Combine(inputDirectory, $"Sample{i}.docx");
                    sampleDoc.Save(samplePath, SaveFormat.Docx);
                }
            }

            // Get all DOCX files in the input directory.
            string[] docxFiles = Directory.GetFiles(inputDirectory, "*.docx");

            foreach (string docxPath in docxFiles)
            {
                // Load the DOCX document.
                Document document = new Document(docxPath);

                // Configure HTML save options with round‑trip information enabled.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportRoundtripInformation = true
                };

                // Determine the output HTML file path.
                string htmlFileName = Path.GetFileNameWithoutExtension(docxPath) + ".html";
                string htmlPath = Path.Combine(outputDirectory, htmlFileName);

                // Save the document as HTML.
                document.Save(htmlPath, htmlOptions);

                // Verify that the HTML file was created.
                if (!File.Exists(htmlPath))
                {
                    throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");
                }
            }

            // Optional: indicate completion.
            Console.WriteLine("Batch conversion completed successfully.");
        }
    }
}
