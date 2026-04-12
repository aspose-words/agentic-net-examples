using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header with a placeholder.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("HeaderPlaceholder");

        // Primary header (used on all other pages) with the same placeholder.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("HeaderPlaceholder");

        // Add three pages of body text.
        builder.MoveToSection(0);
        builder.Writeln("Page 1 body text.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 body text.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 body text.");

        // Set up find-and-replace with a custom callback that decides the replacement
        // based on whether the match is inside the first‑page header.
        FindReplaceOptions options = new FindReplaceOptions(new HeaderReplacingCallback());

        // Perform the replace. The replacement string passed here is ignored because the
        // callback overwrites it.
        int replaced = doc.Range.Replace("HeaderPlaceholder", "unused", options);

        // Validate that at least one replacement occurred.
        if (replaced == 0)
            throw new InvalidOperationException("No header placeholders were replaced.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderReplacementResult.docx");
        doc.Save(outputPath);
    }

    // Callback that provides different replacement text for the first page header.
    private class HeaderReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the ancestor HeaderFooter node that contains the match.
            Node headerNode = args.MatchNode.GetAncestor(NodeType.HeaderFooter);
            if (headerNode is HeaderFooter headerFooter &&
                headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst)
            {
                args.Replacement = "First Page Header";
            }
            else
            {
                args.Replacement = "Other Pages Header";
            }

            return ReplaceAction.Replace;
        }
    }
}
