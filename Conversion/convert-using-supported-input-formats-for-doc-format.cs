using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., DOCX, RTF, HTML, etc.).
        const string inputPath = "input.docx";

        // Path where the converted DOC (Word 97‑2003) file will be saved.
        const string outputPath = "output.doc";

        // Open the source document as a stream.
        using (FileStream inputStream = File.OpenRead(inputPath))
        {
            // Load the document from the input stream. The constructor automatically detects the format.
            Document document = new Document(inputStream);

            // Create an output stream for the DOC file.
            using (FileStream outputStream = File.Create(outputPath))
            {
                // Save the document in the legacy DOC format.
                document.Save(outputStream, SaveFormat.Doc);
            }
        }
    }
}
