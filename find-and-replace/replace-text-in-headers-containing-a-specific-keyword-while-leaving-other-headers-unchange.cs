using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class HeaderFindReplaceExample
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section header (does NOT contain the keyword).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company Report");
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content of the first section.");

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section header (contains the keyword "Special").
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Special Company Overview");
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content of the second section.");

        // Save the original document (optional, for verification).
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Define the keyword to identify target headers and the text to replace.
        const string keyword = "Special";
        const string oldText = "Company";
        const string newText = "Enterprise";

        int totalReplacements = 0;

        // Iterate through each section's primary header.
        foreach (Section section in doc.Sections)
        {
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
                continue;

            // Check if the header contains the specific keyword.
            if (header.Range.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                // Perform replacement only within this header.
                int replaced = header.Range.Replace(oldText, newText, new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in headers containing the keyword.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
