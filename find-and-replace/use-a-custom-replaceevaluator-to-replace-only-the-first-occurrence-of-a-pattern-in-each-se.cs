using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections, each containing three occurrences of the pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int sectionIndex = 1; sectionIndex <= 2; sectionIndex++)
        {
            if (sectionIndex > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage); // Start a new section.

            builder.Writeln($"--- Section {sectionIndex} start ---");

            for (int occ = 1; occ <= 3; occ++)
            {
                builder.Writeln($"This is a PLACEHOLDER in section {sectionIndex}, occurrence {occ}.");
            }

            builder.Writeln($"--- Section {sectionIndex} end ---");
        }

        // Save the input document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Set up the replace callback that replaces only the first match per section.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FirstOccurrencePerSectionReplacer()
        };

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that a replacement occurred in each section (2 sections in this example).
        if (replacedCount != 2)
            throw new InvalidOperationException($"Expected 2 replacements (one per section), but got {replacedCount}.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loadedDoc.Save(outputPath);

        // Output simple confirmation (no interactive input required).
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Input document: {inputPath}");
        Console.WriteLine($"Output document: {outputPath}");
    }

    // Callback that replaces only the first occurrence of the pattern in each section.
    private class FirstOccurrencePerSectionReplacer : IReplacingCallback
    {
        private readonly HashSet<Section> _sectionsReplaced = new HashSet<Section>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine the section that contains the current match.
            Node matchNode = args.MatchNode;
            Section section = matchNode.GetAncestor(NodeType.Section) as Section;

            // Safety check – if we cannot locate a section, skip this match.
            if (section == null)
                return ReplaceAction.Skip;

            // If this section has not been processed yet, replace the match.
            if (!_sectionsReplaced.Contains(section))
            {
                _sectionsReplaced.Add(section);
                args.Replacement = "REPLACED";
                return ReplaceAction.Replace;
            }

            // Otherwise, skip further matches in this section.
            return ReplaceAction.Skip;
        }
    }
}
