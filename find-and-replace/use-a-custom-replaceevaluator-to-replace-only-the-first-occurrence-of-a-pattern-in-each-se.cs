using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections, each containing multiple occurrences of the word "Hello".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int sec = 1; sec <= 3; sec++)
        {
            builder.Writeln($"Section {sec} start.");
            builder.Writeln("Hello world! This is a Hello test.");
            builder.Writeln("Another line without the keyword.");
            builder.Writeln("Hello again in the same section.");
            if (sec < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FirstOccurrencePerSectionReplacer()
        };

        // Replace the first occurrence of "Hello" in each section with "Hi".
        int replacedCount = doc.Range.Replace(new Regex(@"\bHello\b"), "Hi", options);

        // Validate that at least one replacement occurred per section.
        if (replacedCount != 3)
            throw new InvalidOperationException($"Expected 3 replacements, but got {replacedCount}.");

        // Save the modified document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that replaces only the first match in each section.
    private class FirstOccurrencePerSectionReplacer : IReplacingCallback
    {
        private readonly HashSet<Section> _sectionsReplaced = new HashSet<Section>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the section that contains the current match.
            Section? section = args.MatchNode.GetAncestor(NodeType.Section) as Section;
            if (section == null)
                return ReplaceAction.Skip; // Safety check.

            // If this section hasn't been processed yet, replace the match.
            if (_sectionsReplaced.Add(section))
                return ReplaceAction.Replace;

            // Otherwise, skip this match.
            return ReplaceAction.Skip;
        }
    }
}
