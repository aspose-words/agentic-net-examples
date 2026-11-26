using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a first page header and a primary header.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Body content.
        builder.Writeln("This is the body of the document.");

        // Enable different first page header/footer for the first section.
        builder.CurrentSection.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("Header1");

        // Primary header (used for all other pages).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header1");

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for replacement.
        Document loaded = new Document(inputPath);

        // Set up find-and-replace with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new HeaderReplacingCallback()
        };

        // Perform replacement of the placeholder text "Header1".
        int replacedCount = loaded.Range.Replace("Header1", string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No header replacements were performed.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that replaces header text differently for the first page header.
    private class HeaderReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ascend the node hierarchy to find the containing HeaderFooter.
            Node? node = args.MatchNode;
            while (node != null && !(node is HeaderFooter))
                node = node.ParentNode;

            if (node is HeaderFooter headerFooter)
            {
                // Choose replacement based on the header type.
                if (headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst)
                    args.Replacement = "First Header Updated";
                else
                    args.Replacement = "Other Header Updated";
            }
            else
            {
                // Fallback replacement if not inside a header (should not happen here).
                args.Replacement = args.Match.Value;
            }

            return ReplaceAction.Replace;
        }
    }
}
