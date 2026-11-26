using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FirstOccurrencePerSectionReplacer : IReplacingCallback
{
    private readonly HashSet<Section> _replacedSections = new();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Find the section that contains the current match.
        Section? section = args.MatchNode.GetAncestor(NodeType.Section) as Section;
        if (section == null)
            return ReplaceAction.Skip;

        // If this section has not been replaced yet, allow replacement and mark it.
        if (_replacedSections.Add(section))
            return ReplaceAction.Replace;

        // Otherwise skip this match.
        return ReplaceAction.Skip;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple sections, each containing two occurrences of the pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        int numberOfSections = 3;
        for (int i = 1; i <= numberOfSections; i++)
        {
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln($"Section {i} - first PLACEHOLDER.");
            builder.Writeln($"Section {i} - second PLACEHOLDER.");
        }

        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Set up the custom callback to replace only the first occurrence per section.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new FirstOccurrencePerSectionReplacer();

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that exactly one replacement occurred in each section.
        if (replacedCount != numberOfSections)
            throw new InvalidOperationException($"Expected {numberOfSections} replacements, but got {replacedCount}.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
