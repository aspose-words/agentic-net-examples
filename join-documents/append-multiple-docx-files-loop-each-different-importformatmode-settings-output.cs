using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsAppendAndConvert
{
    class Program
    {
        static void Main()
        {
            // Base directory of the application.
            string baseDir = AppContext.BaseDirectory;

            // Folder that contains the source DOCX files.
            string sourceFolder = Path.Combine(baseDir, "Input");

            // Destination folder for the combined PDF.
            string outputFolder = Path.Combine(baseDir, "Output");
            string outputPdfPath = Path.Combine(outputFolder, "Combined.pdf");

            // Ensure the input and output directories exist.
            Directory.CreateDirectory(sourceFolder);
            Directory.CreateDirectory(outputFolder);

            // If there are no DOCX files, create a few sample documents.
            string[] existingDocs = Directory.GetFiles(sourceFolder, "*.docx");
            if (existingDocs.Length == 0)
            {
                for (int i = 1; i <= 3; i++)
                {
                    Document sample = new Document();
                    DocumentBuilder builder = new DocumentBuilder(sample);
                    builder.Writeln($"Sample document {i}");
                    string samplePath = Path.Combine(sourceFolder, $"Sample{i}.docx");
                    sample.Save(samplePath);
                }
            }

            // Collect all DOCX files from the source folder.
            List<string> docxFiles = new List<string>(Directory.GetFiles(sourceFolder, "*.docx"));

            // Define a list of ImportFormatMode values to apply to each file.
            ImportFormatMode[] importModes = new ImportFormatMode[]
            {
                ImportFormatMode.UseDestinationStyles,
                ImportFormatMode.KeepSourceFormatting,
                ImportFormatMode.KeepDifferentStyles
            };

            // Create a blank destination document.
            Document dstDoc = new Document();

            // Loop through each source file and append it using the corresponding mode.
            for (int i = 0; i < docxFiles.Count; i++)
            {
                Document srcDoc = new Document(docxFiles[i]);
                ImportFormatMode mode = importModes[i % importModes.Length];
                dstDoc.AppendDocument(srcDoc, mode);
            }

            // Optional: configure PDF save options (e.g., compliance level).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the combined document as a PDF.
            dstDoc.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"Combined PDF saved to: {outputPdfPath}");
        }
    }
}
