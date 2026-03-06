using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

class AbbreviationExpander
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample paragraphs that contain abbreviations.
        builder.Writeln("AI is a field of computer science.");
        builder.Writeln("NASA launched a new satellite.");
        builder.Writeln("UN is an international organization.");

        // Map each abbreviation to its full expansion.
        var expansions = new Dictionary<string, string>
        {
            { "AI",   "Artificial Intelligence (AI)" },
            { "NASA", "National Aeronautics and Space Administration (NASA)" },
            { "UN",   "United Nations (UN)" }
        };

        // Replace each abbreviation with its expanded form.
        foreach (var kvp in expansions)
        {
            // Ensure only whole-word matches are replaced.
            FindReplaceOptions options = new FindReplaceOptions
            {
                FindWholeWordsOnly = true
            };

            doc.Range.Replace(kvp.Key, kvp.Value, options);
        }

        // Save the resulting document to disk.
        doc.Save("AbbreviationsExpanded.docx");
    }
}
