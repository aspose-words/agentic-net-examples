using System;
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
        builder.Writeln("First Header");

        // Primary header for all other pages.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Other Header");

        // Return to the main body and add enough text to create three pages.
        builder.MoveToSection(0);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Define a callback that replaces header text based on the header type.
        IReplacingCallback callback = new HeaderReplaceCallback();
        FindReplaceOptions options = new FindReplaceOptions(callback);

        // Replace any occurrence of the word "Header" in the document.
        int replacedCount = doc.Range.Replace(new Regex("Header"), "Header", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No header replacements were performed.");

        // Save the modified document.
        doc.Save("Result.docx");
    }

    // Callback that changes the replacement text depending on the header type.
    private class HeaderReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the HeaderFooter node that contains the match.
            HeaderFooter header = args.MatchNode.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            if (header == null)
                return ReplaceAction.Skip; // Not inside a header/footer.

            // Choose replacement based on the header type.
            switch (header.HeaderFooterType)
            {
                case HeaderFooterType.HeaderFirst:
                    args.Replacement = "New First Header";
                    break;
                case HeaderFooterType.HeaderPrimary:
                    args.Replacement = "New Other Header";
                    break;
                default:
                    // For any other header types, keep the original replacement.
                    break;
            }

            return ReplaceAction.Replace;
        }
    }
}
