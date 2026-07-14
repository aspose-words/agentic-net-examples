using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing dates in MM-DD-YYYY format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample dates: 12-31-2020, 01-01-2021, 07-04-2022.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Regular expression to match dates of the form MM-DD-YYYY.
        Regex datePattern = new Regex(@"\b(\d{2})-(\d{2})-(\d{4})\b");

        // Replacement string reorders the captured groups to YYYY-MM-DD.
        string replacement = "$3-$1-$2";

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(datePattern, replacement, new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No dates were replaced. Expected at least one match.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: display the result in the console.
        Console.WriteLine("Original text:");
        Console.WriteLine(doc.GetText().Trim());
        Console.WriteLine();
        Console.WriteLine("Modified text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}
