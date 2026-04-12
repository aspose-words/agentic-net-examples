using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // URL of a sample PDF file.
        const string pdfUrl = "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf";

        // Download the PDF into a memory stream.
        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(pdfUrl).Result;
            response.EnsureSuccessStatusCode();

            byte[] pdfBytes = response.Content.ReadAsByteArrayAsync().Result;
            using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
            {
                // Ensure the stream is positioned at the beginning before loading.
                pdfStream.Position = 0;

                // Load the PDF directly from the stream using PdfLoadOptions.
                Document pdfDocument = new Document(pdfStream, new PdfLoadOptions());

                // Define the output DOCX path.
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Converted.docx");

                // Save the document as DOCX.
                pdfDocument.Save(outputPath, SaveFormat.Docx);

                // Validate that the output file was created and is not empty.
                if (!File.Exists(outputPath))
                {
                    throw new FileNotFoundException("The DOCX file was not created.", outputPath);
                }

                FileInfo info = new FileInfo(outputPath);
                if (info.Length == 0)
                {
                    throw new InvalidOperationException("The DOCX file is empty after conversion.");
                }

                // Optionally, verify the content by loading it back.
                Document verifyDoc = new Document(outputPath);
                Console.WriteLine("Conversion successful. Document text preview:");
                Console.WriteLine(verifyDoc.GetText().Trim());
            }
        }
    }
}
