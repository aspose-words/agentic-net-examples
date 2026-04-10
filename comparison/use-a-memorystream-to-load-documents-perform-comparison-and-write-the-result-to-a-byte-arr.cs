using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("Hello world!");

        // Create the edited document with a different paragraph.
        Document editedDoc = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(editedDoc);
        builderEdited.Writeln("Hello universe!");

        // Save both documents to memory streams.
        using (MemoryStream originalStream = new MemoryStream())
        using (MemoryStream editedStream = new MemoryStream())
        {
            originalDoc.Save(originalStream, SaveFormat.Docx);
            editedDoc.Save(editedStream, SaveFormat.Docx);

            // Reset stream positions before loading.
            originalStream.Position = 0;
            editedStream.Position = 0;

            // Load the documents back from the streams.
            Document loadedOriginal = new Document(originalStream);
            Document loadedEdited = new Document(editedStream);

            // Perform comparison. The original document will receive revisions.
            loadedOriginal.Compare(loadedEdited, "Comparer", DateTime.Now);

            // Verify that at least one revision was created.
            if (loadedOriginal.Revisions.Count == 0)
            {
                throw new InvalidOperationException("No revisions were detected after comparison.");
            }

            // Save the comparison result to a new memory stream.
            using (MemoryStream resultStream = new MemoryStream())
            {
                loadedOriginal.Save(resultStream, SaveFormat.Docx);
                byte[] resultBytes = resultStream.ToArray();

                // Example usage of the resulting byte array (e.g., output its size).
                Console.WriteLine($"Comparison result saved to byte array of length {resultBytes.Length}.");
            }
        }
    }
}
