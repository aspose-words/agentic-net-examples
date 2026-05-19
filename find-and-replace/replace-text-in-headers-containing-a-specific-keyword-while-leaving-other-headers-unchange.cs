using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class HeaderReplaceExample
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderInput.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderOutput.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with two sections, each having a header.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section header – contains the keyword "Confidential".
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Company Report - Confidential");
        builder.MoveToDocumentEnd();

        // First section body.
        builder.Writeln("This is the body of the first section.");

        // Add a second section by inserting a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section header – does NOT contain the keyword.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Company Report - Public");
        builder.MoveToDocumentEnd();

        // Second section body.
        builder.Writeln("This is the body of the second section.");

        // Save the sample document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and replace text only in headers that contain the keyword.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);
        const string keyword = "Confidential";
        const string oldText = "Confidential";
        const string newText = "REDACTED";

        int totalReplacements = 0;

        // Iterate through all sections and their primary headers.
        foreach (Section section in loadedDoc.Sections)
        {
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null && header.Range.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                int replaced = header.Range.Replace(oldText, newText, new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement was performed.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one header replacement, but none occurred.");

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}
