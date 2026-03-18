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
            // Folder that contains the source DOCX files.
            string sourceFolder = @"C:\Docs\Input";

            // Destination file path for the combined PDF.
            string outputPdfPath = @"C:\Docs\Output\Combined.pdf";

            // Collect all DOCX files from the source folder.
            List<string> docxFiles = new List<string>(Directory.GetFiles(sourceFolder, "*.docx"));

            // Define a list of ImportFormatMode values to apply to each file.
            // For demonstration we cycle through the three available modes.
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
                // Load the source document.
                Document srcDoc = new Document(docxFiles[i]);

                // Choose the import mode for this iteration.
                // If there are more files than modes, wrap around the mode array.
                ImportFormatMode mode = importModes[i % importModes.Length];

                // Append the source document to the destination document.
                // The AppendDocument method respects the chosen ImportFormatMode.
                dstDoc.AppendDocument(srcDoc, mode);
            }

            // Optional: configure PDF save options (e.g., compliance level).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set PDF/A-1b compliance.
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the combined document as a PDF.
            dstDoc.Save(outputPdfPath, pdfOptions);
        }
    }
}
