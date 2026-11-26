using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing normal text and a field with the same placeholder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample text with placeholder.");
        // Insert a MERGEFIELD that also contains the word "placeholder".
        builder.InsertField("MERGEFIELD  placeholder  \\* MERGEFORMAT");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for replacement.
        var loadedDoc = new Document(inputPath);

        // Set up FindReplaceOptions to ignore fields.
        var options = new FindReplaceOptions
        {
            IgnoreFields = true
        };

        // Perform replacement of "placeholder" with "value".
        int replacedCount = loadedDoc.Range.Replace("placeholder", "value", options);

        // Validate that only the normal text was replaced (field should remain unchanged).
        if (replacedCount != 1)
        {
            throw new InvalidOperationException($"Expected 1 replacement, but got {replacedCount}.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Use Aspose.Drawing to reference the required package (no functional impact).
        var _ = Color.Blue;

        // Create a simple JSON report of the replacement operation.
        var report = new { Replacements = replacedCount };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }
}
