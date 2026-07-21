using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content.");

        // Save the DOC document to a memory stream (no file is written).
        using (MemoryStream docStream = new MemoryStream())
        {
            sourceDoc.Save(docStream, SaveFormat.Doc);
            byte[] docBytes = docStream.ToArray();

            // Load a new Document from the byte array.
            using (MemoryStream loadStream = new MemoryStream(docBytes))
            {
                Document loadedDoc = new Document(loadStream);

                // Convert the loaded DOC to PDF and write the result to another memory stream.
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    loadedDoc.Save(pdfStream, SaveFormat.Pdf);

                    // Verify that the PDF stream contains data.
                    if (pdfStream.Length == 0)
                        throw new InvalidOperationException("PDF conversion failed: the output stream is empty.");

                    // Optionally write the PDF to a file for inspection (final output, not intermediate).
                    string outputPath = "output.pdf";
                    File.WriteAllBytes(outputPath, pdfStream.ToArray());

                    // Validate that the PDF file was created.
                    if (!File.Exists(outputPath))
                        throw new InvalidOperationException("Expected output PDF was not created.");
                }
            }
        }
    }
}
