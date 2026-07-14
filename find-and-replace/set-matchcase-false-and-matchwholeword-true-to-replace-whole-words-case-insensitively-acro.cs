using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with varied casing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple banana apple Banana APPLE.");

        // Save the source document.
        const string inputFile = "input.docx";
        doc.Save(inputFile);

        // Load the document for processing.
        Document loadedDoc = new Document(inputFile);

        // Configure find‑replace options:
        // - MatchCase = false  => case‑insensitive search.
        // - FindWholeWordsOnly = true => replace only whole words.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Replace the word "apple" with "orange" using the options above.
        int replaceCount = loadedDoc.Range.Replace("apple", "orange", options);

        // Ensure that at least one replacement occurred.
        if (replaceCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputFile = "output.docx";
        loadedDoc.Save(outputFile);
    }
}
