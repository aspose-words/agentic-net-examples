using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

public class ReplaceHeaderCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Determine the header/footer that contains the match.
        var headerFooter = args.MatchNode?.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
        if (headerFooter != null && headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst)
        {
            args.Replacement = "First Header";
        }
        else
        {
            args.Replacement = "Other Header";
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with a first page header and a primary header.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Enable different first page header/footer.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("Header");

        // Primary header (used on other pages).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header");

        // Add enough body content to generate multiple pages.
        builder.MoveToDocumentEnd();
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln($"Body line {i + 1}");
        }

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        var loaded = new Document(inputPath);

        // Set up the replace callback to handle first page header differently.
        var callback = new ReplaceHeaderCallback();
        var options = new FindReplaceOptions { ReplacingCallback = callback };

        // Perform the replacement. The replacement text is supplied by the callback.
        int replacedCount = loaded.Range.Replace("Header", string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
