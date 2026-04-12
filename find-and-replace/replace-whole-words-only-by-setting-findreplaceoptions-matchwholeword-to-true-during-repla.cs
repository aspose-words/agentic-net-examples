using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing the target word both as a whole word and as part of another word.
        builder.Writeln("Jackson will meet you in Jacksonville.");
        builder.Writeln("Jackson is a common name.");

        // Configure find-and-replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacementCount = doc.Range.Replace("Jackson", "Louis", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
        {
            throw new InvalidOperationException("No replacements were made. The FindWholeWordsOnly option may not be working as expected.");
        }

        // Save the modified document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replacements made: {replacementCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
