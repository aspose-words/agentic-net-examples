using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class FirstPageHeaderReplacer : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Find the HeaderFooter that contains the matched node.
        HeaderFooter headerFooter = args.MatchNode.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;

        // If the node is inside a header, decide which replacement to use.
        if (headerFooter != null && headerFooter.IsHeader)
        {
            // Check whether this header belongs to the first page of its section.
            bool isFirstPageHeader = headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst &&
                                     headerFooter.ParentSection.PageSetup.DifferentFirstPageHeaderFooter;

            // Apply a different replacement for the first‑page header.
            args.Replacement = isFirstPageHeader
                ? "First page header replacement text"
                : "Other pages header replacement text";
        }
        else
        {
            // Not a header – keep the original text.
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Ensure each section is configured to use a distinct first‑page header/footer.
        foreach (Section section in doc.Sections)
            section.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Set up find‑and‑replace with a custom callback that distinguishes first‑page headers.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FirstPageHeaderReplacer()
        };

        // Replace the placeholder text "HeaderPlaceholder" wherever it appears in headers.
        doc.Range.Replace(new Regex("HeaderPlaceholder"), "", options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
