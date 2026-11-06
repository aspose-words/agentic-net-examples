using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string folder = Path.Combine(Path.GetTempPath(), "AsposeFindReplaceExample");
        Directory.CreateDirectory(folder);

        string inputPath = Path.Combine(folder, "input.docx");
        string outputPath = Path.Combine(folder, "output.docx");

        // -----------------------------------------------------------------
        // Build a sample document that contains different types of headers.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – enable a different first‑page header/footer.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // First‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
        builder.Writeln("PLACEHOLDER");

        // Primary header for the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("PLACEHOLDER");

        // Return to the main story before inserting a section break.
        builder.MoveToDocumentEnd();

        // Add a second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // In the new section, disable a different first‑page header/footer.
        builder.PageSetup.DifferentFirstPageHeaderFooter = false;

        // Primary header for the second section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("PLACEHOLDER");

        // Save the source document.
        doc.Save(inputPath);

        // --------------------------------------------------------------
        // Load the document and replace the placeholder text with a
        // custom value that depends on the header type.
        // --------------------------------------------------------------
        Document loaded = new Document(inputPath);

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new HeaderReplacingCallback()
        };

        int replacedCount = loaded.Range.Replace("PLACEHOLDER", string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");

        // Save the modified document.
        loaded.Save(outputPath);
    }

    // Callback that substitutes the placeholder with a different string
    // depending on whether the match is inside a first‑page header or any other header.
    private class HeaderReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ascend the node hierarchy until we find the containing HeaderFooter node.
            Node node = args.MatchNode;
            while (node != null && !(node is HeaderFooter))
                node = node.ParentNode;

            // Default replacement text for normal headers.
            string replacementText = "Other Header";

            // If the match is inside a first‑page header, use a special text.
            if (node is HeaderFooter headerFooter &&
                headerFooter.HeaderFooterType == HeaderFooterType.HeaderFirst)
            {
                replacementText = "First Page Header";
            }

            // Set the replacement string.
            args.Replacement = replacementText;
            return ReplaceAction.Replace;
        }
    }
}
