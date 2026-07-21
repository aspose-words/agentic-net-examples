using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample document with some initial text.
        Document initialDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(initialDoc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // 2. Save the document to a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            initialDoc.Save(stream, SaveFormat.Docx);
            stream.Position = 0; // Reset stream for reading.

            // 3. Load the document from the stream.
            Document doc = new Document(stream);

            // 4. Enable track changes.
            doc.StartTrackRevisions("Sample Author", DateTime.Now);

            // 5. Perform a find-and-replace operation while tracking is enabled.
            // Replace the word "fox" with "cat".
            FindReplaceOptions replaceOptions = new FindReplaceOptions();
            doc.Range.Replace("fox", "cat", replaceOptions);

            // 6. Stop tracking changes.
            doc.StopTrackRevisions();

            // 7. List all generated revisions.
            Console.WriteLine("Revisions generated:");
            foreach (Revision rev in doc.Revisions)
            {
                string text = rev.ParentNode != null ? rev.ParentNode.GetText().Trim() : "<no text>";
                Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{text}\"");
            }

            // 8. Save the final document to a file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TrackedDocument.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
