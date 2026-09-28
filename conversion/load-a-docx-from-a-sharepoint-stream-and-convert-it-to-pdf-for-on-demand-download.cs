using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file that will act as the SharePoint document.
        const string inputPath = "sample.docx";
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document stored in SharePoint.");
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Verify the input file was created.
        if (!File.Exists(inputPath))
            throw new InvalidOperationException("The sample DOCX file was not created.");

        // Step 2: Simulate obtaining a SharePoint stream for the DOCX file.
        using MemoryStream sharePointStream = new MemoryStream();
        using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            fileStream.CopyTo(sharePointStream);
        }
        // Reset the stream position before loading.
        sharePointStream.Position = 0;

        // Step 3: Load the document from the simulated SharePoint stream.
        Document docFromSharePoint = new Document(sharePointStream);

        // Step 4: Convert the document to PDF and write it to a simulated response stream.
        using MemoryStream responseStream = new MemoryStream();
        docFromSharePoint.Save(responseStream, SaveFormat.Pdf);

        // Validate that PDF data was written to the response stream.
        if (responseStream.Length == 0)
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");

        // Optional: Save the PDF to a file for verification purposes.
        const string outputPath = "output.pdf";
        responseStream.Position = 0;
        using (FileStream outputFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            responseStream.CopyTo(outputFile);
        }

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output PDF file was not created.");

        // Cleanup: delete temporary files (optional).
        File.Delete(inputPath);
        File.Delete(outputPath);
    }
}
