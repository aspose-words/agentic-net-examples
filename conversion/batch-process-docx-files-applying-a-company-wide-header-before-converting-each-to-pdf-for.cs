using System;
using System.IO;
using Aspose.Words;

public class BatchDocxToPdfWithHeader
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDFs.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputPdfs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files in the input folder.
        for (int i = 1; i <= 3; i++)
        {
            string docxPath = Path.Combine(inputFolder, $"SampleDocument{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is the content of sample document {i}.");
            sampleDoc.Save(docxPath, SaveFormat.Docx);
        }

        // Define the company-wide header text.
        const string headerText = "Company Confidential";

        // Process each DOCX file: add header and convert to PDF.
        foreach (string docxFilePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxFilePath);

            // Add a primary header with the company text.
            DocumentBuilder headerBuilder = new DocumentBuilder(doc);
            headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            headerBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            headerBuilder.Writeln(headerText);

            // Determine the output PDF path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxFilePath);
            string pdfPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.pdf");

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
            {
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
            }
        }
    }
}
