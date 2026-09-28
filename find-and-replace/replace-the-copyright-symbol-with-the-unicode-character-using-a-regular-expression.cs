using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with the text "(c)" that we want to replace.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document. (c) 2026 Company.");
        // Save the original document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform the replacement.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the literal "(c)" (case‑insensitive).
        Regex regex = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Replace each match with the Unicode copyright character.
        int replacedCount = loaded.Range.Replace(regex, "©", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No occurrences of the pattern were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Indicate success (optional, not required for non‑interactive execution).
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
