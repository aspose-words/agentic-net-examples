using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with different first page header.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable different first page header.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header placeholder.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("FirstHeaderPlaceholder");

        // Primary (other pages) header placeholder.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("OtherHeaderPlaceholder");

        // Add body content spanning three pages.
        builder.MoveToSection(0);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for replacement.
        Document loaded = new Document(inputPath);

        // Callback that decides replacement based on header type.
        var callback = new HeaderReplaceCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Replace both placeholders with appropriate text.
        Regex regex = new Regex("(FirstHeaderPlaceholder|OtherHeaderPlaceholder)");
        int replaced = loaded.Range.Replace(regex, string.Empty, options);

        if (replaced == 0)
            throw new InvalidOperationException("No header placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback implementation that checks the header/footer type of the match.
    private class HeaderReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the containing HeaderFooter node, if any.
            HeaderFooter header = args.MatchNode.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;

            if (header != null && header.HeaderFooterType == HeaderFooterType.HeaderFirst)
                args.Replacement = "First Header Updated";
            else
                args.Replacement = "Other Header Updated";

            return ReplaceAction.Replace;
        }
    }
}
