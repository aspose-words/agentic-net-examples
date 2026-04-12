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

        // Add a header with an old copyright year.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("(C) 2006 Aspose Pty Ltd.");

        // Add a footer with the same old copyright year.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("(C) 2006 Aspose Pty Ltd.");

        // Prepare the replacement text using the current year.
        int currentYear = DateTime.Now.Year;
        string oldText = "(C) 2006 Aspose Pty Ltd.";
        string newText = $"(C) {currentYear} Aspose Pty Ltd.";

        // Configure find‑replace options (case‑insensitive, not whole‑word only).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Perform the replacement in every header and footer of each section.
        int totalReplacements = 0;
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                if (headerFooter == null) continue;
                int count = headerFooter.Range.Replace(oldText, newText, options);
                totalReplacements += count;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No copyright year replacements were made.");

        // Save the modified document.
        doc.Save("UpdatedDocument.docx");
    }
}
