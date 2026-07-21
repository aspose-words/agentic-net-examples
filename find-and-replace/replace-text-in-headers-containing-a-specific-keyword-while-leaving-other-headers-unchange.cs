using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class HeaderReplaceExample
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section with a primary header that contains the keyword "Special".
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Special Header - _Company_");
        builder.MoveToDocumentEnd();

        // Add a second section with a primary header that does NOT contain the keyword.
        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        // After the break the builder is positioned in the new section's body.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Regular Header - _Company_");
        builder.MoveToDocumentEnd();

        // Save the original document (optional, for inspection).
        const string inputPath = "HeaderReplaceInput.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        const string keyword = "Special";
        const string placeholder = "_Company_";
        const string replacement = "Acme Corp";

        int totalReplacements = 0;

        // Iterate through all headers in all sections.
        foreach (Section section in loaded.Sections)
        {
            foreach (HeaderFooter header in section.HeadersFooters)
            {
                // Process only primary headers that contain the keyword.
                if (header.HeaderFooterType == HeaderFooterType.HeaderPrimary &&
                    header.GetText().Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    int replaced = header.Range.Replace(placeholder, replacement, new FindReplaceOptions());
                    totalReplacements += replaced;
                }
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were made in the targeted headers.");

        // Save the modified document.
        const string outputPath = "HeaderReplaceOutput.docx";
        loaded.Save(outputPath);
    }
}
