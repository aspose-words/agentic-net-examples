using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple DOC document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content loaded from a byte array.");

        // Save the document to a memory stream in DOC format.
        using (MemoryStream docStream = new MemoryStream())
        {
            source.Save(docStream, SaveFormat.Doc);
            // The stream now contains the DOC file bytes.
            byte[] docBytes = docStream.ToArray();

            // Load a new Document from the byte array.
            using (MemoryStream loadStream = new MemoryStream(docBytes))
            {
                Document loadedDoc = new Document(loadStream);

                // Convert the loaded document to PDF, writing directly to a memory stream.
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    loadedDoc.Save(pdfStream, SaveFormat.Pdf);

                    // Validate that PDF data was written.
                    if (pdfStream.Length == 0)
                        throw new InvalidOperationException("PDF conversion produced an empty stream.");

                    // Optionally write the PDF to a file for verification.
                    File.WriteAllBytes("output.pdf", pdfStream.ToArray());

                    // Verify that the output file was created.
                    if (!File.Exists("output.pdf"))
                        throw new InvalidOperationException("Expected output PDF file was not created.");
                }
            }
        }
    }
}
