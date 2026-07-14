using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Write("First Header");

        // Primary (other pages) header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Other Header");

        // Add body content with page breaks to generate multiple pages.
        builder.MoveToSection(0);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Set up a find-and-replace operation that will replace header text differently
        // depending on whether it belongs to the first page header or a regular header.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new HeaderReplacer()
        };

        // Match the whole header text (either "First Header" or "Other Header").
        Regex headerPattern = new Regex("(First Header|Other Header)", RegexOptions.None);

        int replacedCount = doc.Range.Replace(headerPattern, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No header text was replaced.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }

    // Callback that determines which header is being processed and sets the appropriate replacement text.
    private class HeaderReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the HeaderFooter node that contains the match.
            HeaderFooter headerFooter = args.MatchNode.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            if (headerFooter == null)
                return ReplaceAction.Skip;

            // Choose replacement based on header type.
            switch (headerFooter.HeaderFooterType)
            {
                case HeaderFooterType.HeaderFirst:
                    args.Replacement = "New First Header";
                    break;
                case HeaderFooterType.HeaderPrimary:
                    args.Replacement = "New Header";
                    break;
                default:
                    // For any other header/footer types, keep the original text.
                    args.Replacement = args.Match.Value;
                    break;
            }

            return ReplaceAction.Replace;
        }
    }
}
