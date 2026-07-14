using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Alpha");

        // Save the original document to a memory stream.
        using (MemoryStream msOriginal = new MemoryStream())
        {
            original.Save(msOriginal, SaveFormat.Docx);
            msOriginal.Position = 0;

            // Create the revised document with different content.
            Document revised = new Document();
            DocumentBuilder builderRevised = new DocumentBuilder(revised);
            builderRevised.Writeln("Beta");

            // Save the revised document to a second memory stream.
            using (MemoryStream msRevised = new MemoryStream())
            {
                revised.Save(msRevised, SaveFormat.Docx);
                msRevised.Position = 0;

                // Load both documents from their respective streams.
                Document docOriginal = new Document(msOriginal);
                Document docRevised = new Document(msRevised);

                // Perform the comparison. Revisions will be added to docOriginal.
                docOriginal.Compare(docRevised, "Author", DateTime.Now);

                // Verify that at least one revision was created.
                if (docOriginal.Revisions.Count == 0)
                    throw new InvalidOperationException("Expected at least one revision after comparison.");

                // Save the comparison result to a memory stream.
                using (MemoryStream msResult = new MemoryStream())
                {
                    docOriginal.Save(msResult, SaveFormat.Docx);
                    // Convert the stream to a byte array.
                    byte[] resultBytes = msResult.ToArray();

                    // Example usage of the byte array (e.g., write its length to the console).
                    Console.WriteLine($"Comparison result byte array length: {resultBytes.Length}");
                }
            }
        }
    }
}
