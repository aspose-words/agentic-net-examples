using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that contains whole words and substrings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Jackson will meet you in Jacksonville.");
        builder.Writeln("Jackson is a common name.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Save the document to a local file so we can demonstrate loading it later.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the word "Jackson" with "Louis" using the whole-word option.
        int replacedCount = loaded.Range.Replace("Jackson", "Louis", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole-word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
