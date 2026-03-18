using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a sample document with two sections containing the target word multiple times.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 - first occurrence: TARGET");
        builder.Writeln("Section 1 - second occurrence: TARGET");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - first occurrence: TARGET");
        builder.Writeln("Section 2 - second occurrence: TARGET");

        // Set up find/replace options with a custom callback that limits replacement to the first match per section.
        FindReplaceOptions options = new FindReplaceOptions(new FirstMatchPerSectionCallback());

        // Replace the word "TARGET" with "REPLACED" only for the first occurrence in each section.
        doc.Range.Replace(new Regex(@"\bTARGET\b"), "", options);

        // Save the modified document.
        doc.Save("FirstMatchPerSection.docx");
    }

    // Custom callback implementing IReplacingCallback.
    // It replaces only the first match found in each section.
    private class FirstMatchPerSectionCallback : IReplacingCallback
    {
        // Tracks sections that have already had a replacement performed.
        private readonly HashSet<Section> _replacedSections = new HashSet<Section>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine the section that contains the current match.
            Section section = (Section)args.MatchNode.GetAncestor(NodeType.Section);

            // If this section hasn't been processed yet, perform the replacement.
            if (section != null && !_replacedSections.Contains(section))
            {
                _replacedSections.Add(section);
                args.Replacement = "REPLACED";
                return ReplaceAction.Replace;
            }

            // Skip all subsequent matches in the same section.
            return ReplaceAction.Skip;
        }
    }
}
