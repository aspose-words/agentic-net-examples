using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a sample DOCX file that will act as the document stored in SharePoint.
        const string sampleDocPath = "SampleDocument.docx";
        CreateSampleDocx(sampleDocPath);

        // Simulate obtaining a stream from SharePoint.
        using (FileStream sharePointStream = File.OpenRead(sampleDocPath))
        {
            // Load the document from the SharePoint stream.
            Document document = new Document(sharePointStream);

            // Convert the document to PDF and write the result to a memory stream
            // (simulating an on‑demand download response).
            using (MemoryStream pdfStream = new MemoryStream())
            {
                document.Save(pdfStream, SaveFormat.Pdf);

                // Ensure the stream is ready for reading.
                pdfStream.Position = 0;

                // Validate that the conversion produced data.
                if (pdfStream.Length == 0)
                    throw new InvalidOperationException("PDF conversion resulted in an empty stream.");

                // Optionally, save the PDF to a file for verification.
                const string outputPdfPath = "ConvertedDocument.pdf";
                using (FileStream file = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write))
                {
                    pdfStream.CopyTo(file);
                }

                Console.WriteLine($"PDF conversion successful. Output saved to '{outputPdfPath}'.");
            }
        }

        // Clean up the temporary DOCX file.
        if (File.Exists(sampleDocPath))
            File.Delete(sampleDocPath);
    }

    private static void CreateSampleDocx(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document created for the SharePoint stream conversion example.");
        doc.Save(path, SaveFormat.Docx);
    }
}
