using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing the target word in different cases.
        builder.Writeln("Apple apple APPLE banana apple.");
        builder.Writeln("An apple a day keeps the doctor away.");
        builder.Writeln("Pineapple is not an apple.");

        // Configure find‑replace options:
        // - MatchCase = false  -> case‑insensitive search.
        // - FindWholeWordsOnly = true -> replace only whole word matches.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement: replace the word "apple" with "orange".
        int replacementCount = doc.Range.Replace("apple", "orange", options);

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document to the local file system.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Optional: output a simple confirmation to the console.
        Console.WriteLine($"Replacements made: {replacementCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
