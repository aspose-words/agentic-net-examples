using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with uppercase words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS is a TEST document.");
        builder.Writeln("ANOTHER LINE WITH UPPERCASE WORDS LIKE EXAMPLE.");
        builder.Writeln("Mixed case WORDS should stay unchanged.");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex to match whole words that are all uppercase (minimum two letters).
        Regex upperCaseRegex = new Regex(@"\b[A-Z]{2,}\b", RegexOptions.Compiled);

        // Set up the replacement callback that converts matches to title case.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new TitleCaseCallback()
        };

        // Perform the replacement using an empty replacement string; the callback supplies the actual text.
        int replacedCount = loaded.Range.Replace(upperCaseRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Write a simple JSON report of the operation.
        var report = new
        {
            InputFile = Path.GetFullPath(inputPath),
            OutputFile = Path.GetFullPath(outputPath),
            ReplacementsMade = replacedCount,
            Timestamp = DateTime.UtcNow
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }

    private class TitleCaseCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string word = args.Match.Value;
            if (string.IsNullOrEmpty(word))
                return ReplaceAction.Skip;

            TextInfo ti = CultureInfo.InvariantCulture.TextInfo;
            string titleCase = ti.ToTitleCase(word.ToLowerInvariant());

            args.Replacement = titleCase;
            return ReplaceAction.Replace;
        }
    }
}
