using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing various cases of the word "hello".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World");
        builder.Writeln("hello world");
        builder.Writeln("HELLO WORLD");
        builder.Writeln("HelloWorld"); // Not a whole word, should stay unchanged.

        // Save the sample input document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options: case‑insensitive and whole‑word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true // Correct property name for whole‑word matching.
        };

        // Replace the whole word "hello" with "hi".
        int replacedCount = loaded.Range.Replace("hello", "hi", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
