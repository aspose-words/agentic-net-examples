using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs that contain abbreviations and acronyms.
        builder.Writeln("AI is transforming many industries.");
        builder.Writeln("NASA launched a new mission.");
        builder.Writeln("The CPU performance has improved.");

        // Define the abbreviation expansions.
        var expansions = new (string Abbr, string Full)[]
        {
            ("AI", "Artificial Intelligence (AI)"),
            ("NASA", "National Aeronautics and Space Administration (NASA)"),
            ("CPU", "Central Processing Unit (CPU)")
        };

        // Configure find‑replace options to match whole words only (case‑insensitive).
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true,
            MatchCase = false
        };

        // Replace each abbreviation with its expanded form.
        foreach (var (abbr, full) in expansions)
        {
            doc.Range.Replace(abbr, full, options);
        }

        // Save the resulting document.
        doc.Save("AbbreviationsExpanded.docx");
    }
}
