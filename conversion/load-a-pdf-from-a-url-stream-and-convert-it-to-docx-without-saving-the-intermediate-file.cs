using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;

public class Program
{
    public static async Task Main()
    {
        // URL of a sample PDF file.
        const string pdfUrl = "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf";

        // Download the PDF into a memory stream.
        using HttpClient httpClient = new HttpClient();
        using Stream downloadStream = await httpClient.GetStreamAsync(pdfUrl);
        using MemoryStream pdfMemory = new MemoryStream();
        await downloadStream.CopyToAsync(pdfMemory);
        pdfMemory.Position = 0; // Reset position before loading.

        // Load the PDF document from the memory stream.
        Document pdfDocument = new Document(pdfMemory);

        // Convert the PDF to DOCX and save to a file.
        const string outputPath = "output.docx";
        pdfDocument.Save(outputPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output DOCX was not created.");
        }
    }
}
