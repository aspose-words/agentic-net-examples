using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocxToPdf
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if the input folder is empty.
        string[] sampleNames = { "Sample1.docx", "Sample2.docx", "Sample3.docx" };
        foreach (string fileName in sampleNames)
        {
            string filePath = Path.Combine(inputFolder, fileName);
            if (!File.Exists(filePath))
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is the content of {Path.GetFileNameWithoutExtension(fileName)}.");
                sampleDoc.Save(filePath, SaveFormat.Docx);
            }
        }

        // Process each DOCX file: add a company‑wide header and convert to PDF.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Add a company‑wide header to every section.
            foreach (Section section in doc.Sections)
            {
                // Retrieve the primary header; create it if it does not exist.
                HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                // Append the header text.
                header.AppendParagraph("Company Confidential – Header");
            }

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Indicate completion.
        Console.WriteLine("Batch processing completed successfully.");
    }
}
