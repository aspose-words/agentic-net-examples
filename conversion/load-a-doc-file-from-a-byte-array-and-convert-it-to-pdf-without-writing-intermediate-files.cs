using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample DOC document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content for conversion.");

        // 2. Save the DOC document to a MemoryStream (no file is written).
        using (MemoryStream docStream = new MemoryStream())
        {
            sourceDoc.Save(docStream, SaveFormat.Doc);
            byte[] docBytes = docStream.ToArray();

            // 3. Load a new Document from the byte array.
            using (MemoryStream loadStream = new MemoryStream(docBytes))
            {
                Document loadedDoc = new Document(loadStream);

                // 4. Convert the loaded document to PDF and save to a file.
                const string pdfPath = "output.pdf";
                loadedDoc.Save(pdfPath, SaveFormat.Pdf);

                // 5. Validate that the PDF file was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException("The PDF conversion failed; output file was not created.");

                // Optional: indicate success.
                Console.WriteLine($"PDF successfully created at '{Path.GetFullPath(pdfPath)}'.");
            }
        }
    }
}
