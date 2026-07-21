using System;
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

        // Primary header (used on all other pages).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Other Header");

        // Add some body content spanning two pages.
        builder.MoveToSection(0);
        builder.Writeln("Content on page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content on page 2.");

        // Set up a find‑replace operation that replaces the word "Header"
        // with a different text depending on whether it is in the first‑page header.
        FindReplaceOptions options = new FindReplaceOptions(new HeaderReplaceCallback());

        int replaced = doc.Range.Replace("Header", string.Empty, options);
        if (replaced == 0)
            throw new InvalidOperationException("No header text was replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that decides the replacement text based on the header type.
    private class HeaderReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the HeaderFooter node that contains the match.
            HeaderFooter headerFooter = args.MatchNode.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            if (headerFooter == null)
                return ReplaceAction.Skip; // Not inside a header/footer.

            // Choose replacement based on the header type.
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
