using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using System.Text.RegularExpressions;

namespace FindAndReplaceExample
{
    // Callback that replaces only the first occurrence of a pattern in each section.
    public class FirstOccurrencePerSectionReplacer : IReplacingCallback
    {
        // Tracks sections that have already performed a replacement.
        private readonly HashSet<Section> _sectionsReplaced = new();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the section that contains the current match.
            Section? section = args.MatchNode?.GetAncestor(NodeType.Section) as Section;
            if (section == null)
                return ReplaceAction.Skip; // Safety check.

            // If this section has not been replaced yet, perform the replacement.
            if (!_sectionsReplaced.Contains(section))
            {
                _sectionsReplaced.Add(section);
                // The replacement text is already supplied by the caller (args.Replacement),
                // so we simply allow the replace operation to proceed.
                return ReplaceAction.Replace;
            }

            // Skip subsequent matches in the same section.
            return ReplaceAction.Skip;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document with multiple sections, each containing several occurrences of the word "pattern".
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Section 1
            builder.Writeln("Section 1 - First occurrence: pattern.");
            builder.Writeln("Section 1 - Second occurrence: pattern.");
            builder.Writeln("Section 1 - Third occurrence: pattern.");

            // Insert a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 2
            builder.Writeln("Section 2 - First occurrence: pattern.");
            builder.Writeln("Section 2 - Second occurrence: pattern.");

            // Insert another new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 3
            builder.Writeln("Section 3 - Only occurrence: pattern.");

            // Save the original document (optional, for inspection).
            doc.Save("input.docx");

            // Set up find-and-replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new FirstOccurrencePerSectionReplacer()
            };

            // Perform the replacement: replace the word "pattern" with "replaced" only on its first appearance per section.
            int replacedCount = doc.Range.Replace("pattern", "replaced", options);

            // Validate that we replaced exactly one occurrence per section (three sections in this example).
            const int expectedReplacements = 3;
            if (replacedCount != expectedReplacements)
                throw new InvalidOperationException($"Expected {expectedReplacements} replacements, but got {replacedCount}.");

            // Save the modified document.
            doc.Save("output.docx");
        }
    }
}
