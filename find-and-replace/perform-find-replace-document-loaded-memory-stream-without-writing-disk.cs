using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document entirely in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello _FullName_! This is a sample document.");

        // Save the original document to a memory stream (no disk I/O).
        using (MemoryStream sourceStream = new MemoryStream())
        {
            doc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset stream for reading.

            // Load the document from the memory stream.
            Document loadedDoc = new Document(sourceStream);

            // Perform a find‑and‑replace operation on the loaded document.
            int replacements = loadedDoc.Range.Replace("_FullName_", "John Doe");
            Console.WriteLine($"Replacements made: {replacements}");

            // Save the modified document to another memory stream.
            using (MemoryStream resultStream = new MemoryStream())
            {
                loadedDoc.Save(resultStream, SaveFormat.Docx);
                resultStream.Position = 0; // Reset for further use.

                // Example: read back the text to verify the replacement.
                Document finalDoc = new Document(resultStream);
                Console.WriteLine(finalDoc.GetText().Trim());
            }
        }
    }
}
