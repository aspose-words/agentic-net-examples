using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs containing the placeholder "(c)".
        builder.Writeln("This is a test (c) for copyright.");
        builder.Writeln("Another line with (c) symbol.");

        // Define a regular expression that matches the literal "(c)" (case‑insensitive).
        Regex pattern = new Regex(@"\(c\)", RegexOptions.IgnoreCase);

        // Replace each match with the Unicode copyright character ©.
        int replacementCount = doc.Range.Replace(pattern, "\u00A9");

        // Validate that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("No occurrences of the pattern were found to replace.");

        // Save the modified document to a local file.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Optional: indicate success (no interactive input required).
        Console.WriteLine($"Replacements made: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
