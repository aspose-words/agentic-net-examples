using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocumentComparer
{
    /// <summary>
    /// Compares two Word documents supplied as byte arrays.
    /// The comparison result (original document with revision marks) is returned as a byte array.
    /// </summary>
    /// <param name="originalBytes">Byte array containing the original document.</param>
    /// <param name="editedBytes">Byte array containing the edited document.</param>
    /// <returns>Byte array of the comparison result saved in DOCX format.</returns>
    public static byte[] Compare(byte[] originalBytes, byte[] editedBytes)
    {
        // Load the original document from a memory stream.
        using (var originalStream = new MemoryStream(originalBytes))
        {
            var docOriginal = new Document(originalStream);

            // Load the edited document from a memory stream.
            using (var editedStream = new MemoryStream(editedBytes))
            {
                var docEdited = new Document(editedStream);

                // Perform the comparison only if both documents have no existing revisions.
                if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
                {
                    // Compare adds revision marks to docOriginal.
                    docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
                }

                // Save the comparison result to a memory stream in DOCX format.
                using (var resultStream = new MemoryStream())
                {
                    docOriginal.Save(resultStream, SaveFormat.Docx);
                    // Return the underlying byte array.
                    return resultStream.ToArray();
                }
            }
        }
    }
}

public class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Simple demonstration – replace the file paths with real files when testing.
        string originalPath = "original.docx";
        string editedPath   = "edited.docx";
        string resultPath   = "comparisonResult.docx";

        if (!File.Exists(originalPath) || !File.Exists(editedPath))
        {
            Console.WriteLine("Please ensure both input files exist.");
            return;
        }

        byte[] original = File.ReadAllBytes(originalPath);
        byte[] edited   = File.ReadAllBytes(editedPath);
        byte[] result   = DocumentComparer.Compare(original, edited);
        File.WriteAllBytes(resultPath, result);
        Console.WriteLine($"Comparison saved to '{resultPath}'.");
    }
}
