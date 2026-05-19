using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a different first‑page header.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a distinct header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("Header Placeholder");

        // Primary (other pages) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header Placeholder");

        // Add three pages of body text.
        builder.MoveToSection(0);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for replacement.
        Document loaded = new Document(inputPath);

        // Callback that replaces the first‑page header differently from other headers.
        var callback = new HeaderReplaceCallback();

        // Use the callback in FindReplaceOptions.
        FindReplaceOptions options = new FindReplaceOptions(callback);

        // Replace the placeholder text in all headers.
        int replaced = loaded.Range.Replace(new Regex("Header Placeholder"), "", options);
        if (replaced == 0)
            throw new InvalidOperationException("No header text was replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback implementation for custom header replacement.
    private class HeaderReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine the HeaderFooter that contains the current match.
            HeaderFooter header = args.MatchNode.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;

            // Choose replacement based on header type.
            if (header != null && header.HeaderFooterType == HeaderFooterType.HeaderFirst)
                args.Replacement = "New First Header";
            else
                args.Replacement = "New Primary Header";

            return ReplaceAction.Replace;
        }
    }
}
