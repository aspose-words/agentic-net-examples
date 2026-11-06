using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with words longer than ten characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(
            "This document contains some extraordinarilylongword and anotherSupercalifragilisticexpialidocious word.");
        doc.Save("input.docx");

        // Regex to match words longer than ten characters.
        Regex longWordRegex = new Regex(@"\b\w{11,}\b", RegexOptions.Compiled);

        // Collect original words that are replaced.
        List<string> replacedWords = new List<string>();

        // Set up FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LongWordReplacer(replacedWords)
        };

        // Perform the replacement using the custom callback.
        // The replacement string is ignored because the callback provides its own replacement.
        int replacedCount = doc.Range.Replace(longWordRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement of a long word.");

        // Save the modified document.
        doc.Save("output.docx");

        // Create a simple JSON report of the replacements.
        var report = new
        {
            ReplacedCount = replacedCount,
            Words = replacedWords
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }

    // Custom callback that records each match and replaces it with asterisks.
    private class LongWordReplacer : IReplacingCallback
    {
        private readonly List<string> _replacedWords;

        public LongWordReplacer(List<string> replacedWords)
        {
            _replacedWords = replacedWords ?? throw new ArgumentNullException(nameof(replacedWords));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string originalWord = args.Match.Value;
            _replacedWords.Add(originalWord);
            args.Replacement = new string('*', originalWord.Length);
            return ReplaceAction.Replace;
        }
    }
}
