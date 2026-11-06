using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("foo is at the start of this paragraph.");
        builder.Writeln("This paragraph contains foo in the middle.");
        builder.Writeln("foo appears again at the beginning.");
        builder.Writeln("No keyword here.");
        builder.Writeln("   foo with leading spaces should not be replaced.");

        // Set up the custom callback that replaces only when 'foo' is at the start of a paragraph.
        var callback = new StartOfParagraphCallback();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("foo", "bar", options);

        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Create a simple JSON report of the replacements.
        var report = new
        {
            Replacements = replacedCount,
            ParagraphsModified = callback.ModifiedParagraphTexts
        };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }

    private class StartOfParagraphCallback : IReplacingCallback
    {
        public List<string> ModifiedParagraphTexts { get; } = new List<string>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match node is expected to be a Run.
            if (args.MatchNode is not Run run)
                return ReplaceAction.Skip;

            // Get the paragraph that contains this run.
            Paragraph paragraph = run.ParentParagraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Verify that the run is the first node in the paragraph.
            if (paragraph.FirstChild != run)
                return ReplaceAction.Skip;

            // Ensure the run's text actually starts with the matched word.
            if (!run.Text.StartsWith(args.Match.Value, StringComparison.Ordinal))
                return ReplaceAction.Skip;

            // Perform the replacement.
            args.Replacement = "bar";

            // Record the paragraph's text for reporting.
            ModifiedParagraphTexts.Add(paragraph.GetText());

            return ReplaceAction.Replace;
        }
    }
}
