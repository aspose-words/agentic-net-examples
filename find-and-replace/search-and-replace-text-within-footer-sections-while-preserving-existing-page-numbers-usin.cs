using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplaceLogger : IReplacingCallback
{
    public List<string> Matches { get; } = new List<string>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Record the matched text.
        Matches.Add(args.Match.Value);
        // Perform the replacement.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content.
        builder.Writeln("This is the main body of the document.");
        builder.Writeln("It has multiple pages.");

        // Insert a footer with placeholder text and page fields.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Report - Confidential - Page ");
        builder.InsertField("PAGE", "?");
        builder.Write(" of ");
        builder.InsertField("NUMPAGES", "?");
        builder.Writeln(); // Ensure footer ends with a line break.

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Set up the replace logger and options.
        ReplaceLogger logger = new ReplaceLogger();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger,
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Perform replacement only in footers.
        int totalReplacements = 0;
        foreach (Section section in loadedDoc.Sections)
        {
            // Primary footer.
            HeaderFooter primaryFooter = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (primaryFooter != null)
                totalReplacements += primaryFooter.Range.Replace("Confidential", "Public", options);

            // First page footer (if any).
            HeaderFooter firstFooter = section.HeadersFooters[HeaderFooterType.FooterFirst];
            if (firstFooter != null)
                totalReplacements += firstFooter.Range.Replace("Confidential", "Public", options);

            // Even pages footer (if any).
            HeaderFooter evenFooter = section.HeadersFooters[HeaderFooterType.FooterEven];
            if (evenFooter != null)
                totalReplacements += evenFooter.Range.Replace("Confidential", "Public", options);
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in footers.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Write a JSON report of the matches that were replaced.
        string jsonReport = JsonConvert.SerializeObject(new
        {
            Replacements = logger.Matches,
            TotalCount = totalReplacements
        }, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}
