using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class FirstOccurrencePerSectionCallback : IReplacingCallback
{
    private readonly HashSet<Section> _replacedSections = new HashSet<Section>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Find the section that contains the current match.
        var matchNode = args.MatchNode;
        if (matchNode == null)
            return ReplaceAction.Skip;

        var section = matchNode.GetAncestor(NodeType.Section) as Section;
        if (section == null)
            return ReplaceAction.Skip;

        // If we have already replaced a match in this section, skip further matches.
        if (_replacedSections.Contains(section))
            return ReplaceAction.Skip;

        // First match in this section – perform the replacement.
        _replacedSections.Add(section);
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with several sections, each containing the word "pattern" multiple times.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Section 1: pattern pattern pattern");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2: pattern pattern");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3: no matching word here");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 4: pattern");

        // Set up find-and-replace options with a custom callback that replaces only the first occurrence per section.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FirstOccurrencePerSectionCallback()
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("pattern", "replaced", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
