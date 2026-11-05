using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for SharePoint stream conversion.");

        // Save the sample DOCX to a temporary file.
        string tempDocxPath = Path.Combine(Path.GetTempPath(), "sample_input.docx");
        sourceDoc.Save(tempDocxPath, SaveFormat.Docx);

        // Simulate obtaining a SharePoint stream by reading the file into a MemoryStream.
        using (FileStream fileStream = new FileStream(tempDocxPath, FileMode.Open, FileAccess.Read))
        using (MemoryStream sharePointStream = new MemoryStream())
        {
            fileStream.CopyTo(sharePointStream);
            sharePointStream.Position = 0; // Reset before loading.

            // Load the DOCX from the simulated SharePoint stream.
            Document docFromStream = new Document(sharePointStream);

            // Convert to PDF and write to a simulated response stream.
            using (MemoryStream responseStream = new MemoryStream())
            {
                docFromStream.Save(responseStream, SaveFormat.Pdf);

                // Validate that PDF data was written.
                if (responseStream.Length == 0)
                {
                    throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
                }

                // Optionally, write the PDF to a file to verify on disk.
                string outputPdfPath = Path.Combine(Path.GetTempPath(), "output.pdf");
                responseStream.Position = 0; // Reset before copying to file.
                using (FileStream outFile = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write))
                {
                    responseStream.CopyTo(outFile);
                }

                if (!File.Exists(outputPdfPath))
                {
                    throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }

        // Clean up temporary files.
        if (File.Exists(tempDocxPath))
        {
            File.Delete(tempDocxPath);
        }
    }
}
