using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocxToPdfExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string inputFile = @"C:\Docs\SampleDocument.docx";

            // Path where the resulting PDF will be saved.
            string outputFile = @"C:\Docs\SampleDocument.pdf";

            // Load the DOCX document from the file system.
            Document doc = new Document(inputFile);

            // Create a SaveOptions object that is appropriate for PDF output.
            // This uses the provided CreateSaveOptions rule.
            SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

            // Optional: enable high‑quality rendering for better visual fidelity.
            // pdfOptions.UseHighQualityRendering = true;

            // Save the document as PDF using the options object.
            // This follows the Save(string, SaveOptions) rule.
            doc.Save(outputFile, pdfOptions);
        }
    }
}
