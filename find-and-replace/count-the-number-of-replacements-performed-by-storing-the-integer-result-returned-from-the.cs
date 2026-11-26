using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that will be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is quick.");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for replacement.
        Document loaded = new Document(inputPath);

        // Perform a plain‑text replacement and capture the number of replacements.
        const string findText = "quick";
        const string replaceText = "swift";
        int replacementCount = loaded.Range.Replace(findText, replaceText, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Write a simple JSON report containing the replacement count.
        var report = new { ReplacementsMade = replacementCount };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);

        // Output the count to the console (non‑interactive).
        Console.WriteLine($"Replacements performed: {replacementCount}");
    }
}
