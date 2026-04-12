using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ProgressReporter
{
    // Reports the number of replacements performed for a given pattern.
    public void Report(string pattern, int count)
    {
        Console.WriteLine($"Pattern \"{pattern}\" was replaced {count} time(s).");
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {Name}, welcome to {Place}.");
        builder.Writeln("Your order {OrderId} is confirmed.");
        builder.Writeln("Thank you, {Name}!");

        // Define the patterns to replace and their corresponding replacements.
        var replacements = new List<(string Pattern, string Replacement)>
        {
            ("{Name}", "Alice"),
            ("{Place}", "Wonderland"),
            ("{OrderId}", "12345")
        };

        // Initialize the progress reporter.
        ProgressReporter reporter = new ProgressReporter();

        // Perform each replacement and report the count.
        foreach (var (pattern, replacement) in replacements)
        {
            int count = doc.Range.Replace(pattern, replacement);
            // Validate that at least one replacement occurred.
            if (count == 0)
                throw new InvalidOperationException($"No occurrences of \"{pattern}\" were found.");
            reporter.Report(pattern, count);
        }

        // Save the modified document to the local file system.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to \"{outputPath}\".");
    }
}
