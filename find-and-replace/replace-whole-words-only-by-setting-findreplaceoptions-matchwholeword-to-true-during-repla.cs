using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package reference

public class Program
{
    public static void Main()
    {
        // Create a sample document with a word that appears both as a whole word and as part of another word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Jackson will meet you in Jacksonville.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("Jackson", "Louis", options);

        // Validate that a replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole-word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Indicate success.
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }
}
