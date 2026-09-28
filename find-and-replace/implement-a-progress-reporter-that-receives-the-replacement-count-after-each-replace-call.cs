using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package reference
using Newtonsoft.Json;        // Required package reference

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {Name}, your order {OrderId} is ready.");
        builder.Writeln("Dear {Name}, please confirm your order {OrderId}.");
        doc.Save("input.docx");

        // Define the find/replace pairs.
        var replacements = new List<(string Find, string Replace)>
        {
            ("{Name}", "John Doe"),
            ("{OrderId}", "12345")
        };

        // Perform each replacement and report progress.
        foreach (var (find, replace) in replacements)
        {
            int replacedCount = doc.Range.Replace(find, replace, new FindReplaceOptions());
            ReportProgress(find, replace, replacedCount);

            if (replacedCount == 0)
                throw new InvalidOperationException($"Expected at least one replacement for '{find}'.");
        }

        // Save the modified document.
        doc.Save("output.docx");
    }

    private static void ReportProgress(string find, string replace, int count)
    {
        Console.WriteLine($"Replaced '{find}' with '{replace}' {count} time(s).");
    }
}
