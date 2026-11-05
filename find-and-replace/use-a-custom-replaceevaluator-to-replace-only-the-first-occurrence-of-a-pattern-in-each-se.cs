using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FirstOccurrencePerSectionReplacer : IReplacingCallback
{
    private readonly HashSet<Section> _replacedSections = new HashSet<Section>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Ensure the match node is not null.
        if (args.MatchNode == null)
            return ReplaceAction.Skip;

        // Find the section that contains the current match.
        Section? section = args.MatchNode.GetAncestor(NodeType.Section) as Section;
        if (section == null)
            return ReplaceAction.Skip;

        // Replace only the first occurrence in each section.
        if (_replacedSections.Contains(section))
            return ReplaceAction.Skip;

        _replacedSections.Add(section);
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document with two sections, each containing multiple occurrences of the pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 start. PLACEHOLDER appears here. Some more text. PLACEHOLDER again.");
        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        // Section 2
        builder.Writeln("Section 2 start. PLACEHOLDER appears here too. Another PLACEHOLDER here.");

        // Set up the replace callback to replace only the first occurrence per section.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FirstOccurrencePerSectionReplacer()
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
