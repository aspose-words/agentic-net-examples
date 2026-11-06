using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document for PDF/A-2u conversion.");

        // Save the sample document to a memory stream (simulating a network stream).
        using MemoryStream networkStream = new MemoryStream();
        sampleDoc.Save(networkStream, SaveFormat.Docx);

        // Reset the stream position before loading.
        networkStream.Position = 0;

        // Load the DOCX from the simulated network stream.
        Document loadedDoc = new Document(networkStream);

        // Configure PDF/A-2u compliance options.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as PDF/A-2u to a file.
        string outputPath = "output.pdf";
        loadedDoc.Save(outputPath, saveOptions);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF/A-2u output file was not created.");
        }

        // Verify that the file size is greater than zero.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The PDF/A-2u output file is empty.");
        }
    }
}
