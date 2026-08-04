using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with words in different cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple apple APPLE banana");
        builder.Writeln("Pineapple is not an apple.");

        // Save the source document (demonstrates the create‑save workflow).
        doc.Save("input.docx");

        // Load the document from the file system (demonstrates the load workflow).
        Document loaded = new Document("input.docx");

        // Configure find‑replace options: case‑insensitive and whole‑word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Replace all whole-word occurrences of "apple" with "orange".
        int replacedCount = loaded.Range.Replace("apple", "orange", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}
