using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // URL of a sample PDF file.
        const string pdfUrl = "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf";

        // Download the PDF into a byte array.
        using (HttpClient httpClient = new HttpClient())
        {
            byte[] pdfBytes = httpClient.GetByteArrayAsync(pdfUrl).GetAwaiter().GetResult();

            // Load the PDF from a memory stream.
            using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
            {
                pdfStream.Position = 0; // Ensure the stream is at the beginning.

                // Create an Aspose.Words Document from the PDF stream.
                Document doc = new Document(pdfStream);

                // Convert and save the document as DOCX.
                const string outputPath = "output.docx";
                doc.Save(outputPath, SaveFormat.Docx);

                // Verify that the DOCX file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("The DOCX output file was not created.");
            }
        }
    }
}
