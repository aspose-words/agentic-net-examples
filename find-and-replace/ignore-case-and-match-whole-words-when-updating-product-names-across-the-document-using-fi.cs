using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document with various occurrences of the word "Apple".
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our product list:");
        builder.Writeln("Apple");          // exact match
        builder.Writeln("apple");          // different case
        builder.Writeln("Apple iPhone");  // part of a phrase (whole word)
        builder.Writeln("Pineapple");     // word containing the pattern
        builder.Writeln("APPLE");          // upper case
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and configure FindReplaceOptions to ignore case
        // and replace only whole-word matches.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // ignore character case
            FindWholeWordsOnly = true   // replace only whole words
        };

        // Perform the replacement: "Apple" -> "Orange".
        int replacedCount = loaded.Range.Replace("Apple", "Orange", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Optional: output the result to the console for verification.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine("Modified document text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}
