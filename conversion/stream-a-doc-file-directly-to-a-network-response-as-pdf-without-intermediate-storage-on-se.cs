using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content for PDF conversion.");
        sourceDoc.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Simulate a network response by writing the PDF directly to a MemoryStream.
        using MemoryStream responseStream = new MemoryStream();
        doc.Save(responseStream, SaveFormat.Pdf);

        // Validate that PDF data was written to the simulated response stream.
        if (responseStream.Length == 0)
        {
            throw new InvalidOperationException("No PDF data was written to the simulated response stream.");
        }

        // (Optional) Write the PDF to a file to demonstrate the result.
        // This does not constitute intermediate storage for the response.
        responseStream.Position = 0;
        using FileStream file = new FileStream("output.pdf", FileMode.Create, FileAccess.Write);
        responseStream.CopyTo(file);
        file.Flush();

        // Verify that the output file was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
