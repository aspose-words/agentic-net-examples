using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToPdfConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\Input.docx";

        // Path where the resulting PDF will be saved.
        string outputFile = @"C:\Docs\Output.pdf";

        // Load the DOCX document using the Document(string) constructor.
        Document doc = new Document(inputFile);

        // Create a PdfSaveOptions instance appropriate for PDF output.
        // The factory method returns a SaveOptions object; cast it to PdfSaveOptions.
        PdfSaveOptions pdfOptions = (PdfSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Fine‑tuned control: enable high‑quality rendering and reduce memory usage.
        pdfOptions.UseHighQualityRendering = true;
        pdfOptions.MemoryOptimization = true;

        // Example of additional PDF‑specific option – limit outline to three heading levels.
        pdfOptions.OutlineOptions.HeadingsOutlineLevels = 3;

        // Save the document as PDF using the overload that accepts SaveOptions.
        doc.Save(outputFile, pdfOptions);
    }
}
