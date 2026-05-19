using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document in memory.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Alpha");

        // Save the original document to a MemoryStream.
        using (MemoryStream msOriginal = new MemoryStream())
        {
            original.Save(msOriginal, SaveFormat.Docx);
            msOriginal.Position = 0; // Reset for reading.

            // Create the revised document in memory.
            Document revised = new Document();
            DocumentBuilder builder2 = new DocumentBuilder(revised);
            builder2.Writeln("Beta");

            // Save the revised document to a MemoryStream.
            using (MemoryStream msRevised = new MemoryStream())
            {
                revised.Save(msRevised, SaveFormat.Docx);
                msRevised.Position = 0; // Reset for reading.

                // Load both documents from their streams.
                Document docOriginal = new Document(msOriginal);
                Document docRevised = new Document(msRevised);

                // Perform comparison.
                docOriginal.Compare(docRevised, "Author", DateTime.Now);

                // Verify that revisions were created.
                if (docOriginal.Revisions.Count == 0)
                    throw new InvalidOperationException("Expected at least one revision after comparison.");

                // Save the comparison result to a byte array.
                using (MemoryStream resultStream = new MemoryStream())
                {
                    docOriginal.Save(resultStream, SaveFormat.Docx);
                    byte[] resultBytes = resultStream.ToArray();

                    // Example usage of the resulting byte array (e.g., write its length).
                    Console.WriteLine($"Comparison result byte array length: {resultBytes.Length}");
                }
            }
        }
    }
}
