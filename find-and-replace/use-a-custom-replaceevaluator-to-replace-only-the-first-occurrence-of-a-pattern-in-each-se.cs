using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections, each containing the pattern "PLACEHOLDER" multiple times.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - first occurrence: PLACEHOLDER");
        builder.Writeln("Section 1 - second occurrence: PLACEHOLDER");

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - first occurrence: PLACEHOLDER");
        builder.Writeln("Section 2 - another occurrence: PLACEHOLDER");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up a custom callback that replaces only the first match in each section.
        var callback = new FirstOccurrencePerSectionReplacer();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement using a regular expression.
        int replacedCount = loaded.Range.Replace(new Regex("PLACEHOLDER"), "REPLACED", options);

        // Validate that at least one replacement occurred in each section (2 sections expected).
        if (replacedCount < 2)
            throw new InvalidOperationException("Expected at least one replacement per section.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that allows replacement only for the first match found in each section.
    private class FirstOccurrencePerSectionReplacer : IReplacingCallback
    {
        private readonly HashSet<Section> _sectionsReplaced = new HashSet<Section>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine the section that contains the current match.
            Node matchNode = args.MatchNode;
            Section section = (Section)matchNode.GetAncestor(NodeType.Section);
            if (section == null)
                return ReplaceAction.Skip;

            // If this section already had a replacement, skip further matches.
            if (_sectionsReplaced.Contains(section))
                return ReplaceAction.Skip;

            // Mark the section as having performed its first replacement and allow it.
            _sectionsReplaced.Add(section);
            return ReplaceAction.Replace;
        }
    }
}
