using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a first section with a primary header that contains the keyword "Confidential".
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Report - Confidential");
        builder.MoveToDocumentEnd();

        // Add a second section with a primary header that does NOT contain the keyword.
        Section secondSection = new Section(doc);
        doc.Sections.Add(secondSection);
        // Move the builder to the newly added section (index 1, zero‑based).
        builder.MoveToSection(1);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Report - Public");
        builder.MoveToDocumentEnd();

        // Save the sample document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define the keyword that determines which headers should be processed.
        const string keyword = "Confidential";

        // Define the text to find and its replacement.
        const string findText = "Report";
        const string replaceText = "Summary";

        int totalReplacements = 0;

        // Iterate through all sections and their headers.
        foreach (Section section in loadedDoc.Sections)
        {
            foreach (HeaderFooter header in section.HeadersFooters)
            {
                if (header == null) continue;

                // Process only headers that contain the keyword (case‑insensitive).
                if (header.Range.Text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    int replaced = header.Range.Replace(findText, replaceText, new FindReplaceOptions());
                    totalReplacements += replaced;
                }
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in headers containing the keyword.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
