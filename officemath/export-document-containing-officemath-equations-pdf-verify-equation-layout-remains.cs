using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportOfficeMathToPdf
{
    static void Main()
    {
        // Input DOCX that contains OfficeMath equations.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMath.docx");

        // Output PDF file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMath.pdf");

        // Load the document.
        Document doc = new Document(inputPath);

        // Ensure the document actually contains OfficeMath objects.
        int officeMathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (officeMathCount == 0)
            throw new InvalidOperationException("The source document does not contain any OfficeMath equations.");

        // Create PDF save options (default settings preserve layout).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF.
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("PDF file was not created.", outputPath);

        long fileSize = new FileInfo(outputPath).Length;
        if (fileSize == 0)
            throw new InvalidOperationException("Generated PDF file is empty; equation layout may be lost.");

        Console.WriteLine($"Exported {officeMathCount} OfficeMath equation(s) to PDF successfully. File size: {fileSize} bytes.");
    }
}
