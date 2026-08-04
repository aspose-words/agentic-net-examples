using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content.");

        // Save the DOC document to a memory stream (no file is written).
        using (MemoryStream docStream = new MemoryStream())
        {
            source.Save(docStream, SaveFormat.Doc);
            // Obtain the byte array representing the DOC file.
            byte[] docBytes = docStream.ToArray();

            // Load a new Document from the byte array.
            using (MemoryStream loadStream = new MemoryStream(docBytes))
            {
                Document loadedDoc = new Document(loadStream);

                // Convert the loaded DOC to PDF and write the result to a memory stream.
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    loadedDoc.Save(pdfStream, SaveFormat.Pdf);

                    // Verify that the PDF data was written.
                    if (pdfStream.Length == 0)
                        throw new InvalidOperationException("No PDF data was written to the output stream.");

                    // Optionally, write the PDF to a file for verification.
                    File.WriteAllBytes("output.pdf", pdfStream.ToArray());

                    // Verify that the output file exists.
                    if (!File.Exists("output.pdf"))
                        throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
