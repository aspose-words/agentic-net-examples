using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello _Name_! This is a sample document.");
        builder.Writeln("Another line with _Name_ placeholder.");

        // Save the document to a memory stream (no disk I/O).
        using (MemoryStream sourceStream = new MemoryStream())
        {
            doc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset for reading.

            // Load the document from the memory stream.
            Document loadedDoc = new Document(sourceStream);

            // Perform a find‑and‑replace operation.
            int replacements = loadedDoc.Range.Replace("_Name_", "World");
            if (replacements == 0)
                throw new InvalidOperationException("Expected at least one replacement, but none were made.");

            // Optionally, verify the replacement result.
            string resultText = loadedDoc.GetText();
            Console.WriteLine("Replaced document text:");
            Console.WriteLine(resultText.Trim());

            // Save the modified document to another memory stream (still no disk I/O).
            using (MemoryStream resultStream = new MemoryStream())
            {
                loadedDoc.Save(resultStream, SaveFormat.Docx);
                // The resultStream now contains the updated document bytes.
                // It can be used further as needed without writing to disk.
            }
        }
    }
}
