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

        // Add sample paragraphs. Some contain the word "Placeholder" which we will replace.
        builder.Writeln("First paragraph with Placeholder.");
        builder.Writeln("Second paragraph without the keyword.");
        builder.Writeln("Third paragraph with Placeholder.");

        // Perform find-and-replace.
        // The replacement string uses the meta‑character "&p" to insert a paragraph break
        // after each occurrence of "Placeholder".
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementCount = doc.Range.Replace("Placeholder", "Placeholder&p", options);

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No occurrences of the search text were found.");

        // Save the modified document to the local file system.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
