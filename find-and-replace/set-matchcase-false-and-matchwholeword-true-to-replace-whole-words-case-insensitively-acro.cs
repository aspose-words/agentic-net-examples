using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly but ensures reference

public class Program
{
    public static void Main()
    {
        // Create a sample document with varied casing of the word "hello".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World. This is a test.");
        builder.Writeln("hello world appears again.");
        builder.Writeln("HELLO WORLD in uppercase.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options: case‑insensitive, whole‑word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // Ignore case.
            FindWholeWordsOnly = true   // Replace only whole words.
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("hello", "Hi", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
