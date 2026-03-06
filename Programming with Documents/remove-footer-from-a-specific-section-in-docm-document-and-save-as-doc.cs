using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\source.docm";

        // Path where the resulting DOCX will be saved.
        string outputPath = @"C:\Docs\result.docx";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Index of the section from which the footer should be removed (0‑based).
        int targetSectionIndex = 1; // Example: second section.

        // Ensure the requested section exists.
        if (targetSectionIndex >= 0 && targetSectionIndex < doc.Sections.Count)
        {
            Section section = doc.Sections[targetSectionIndex];

            // Remove all types of footers from this section.
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            footer?.Remove();

            footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
            footer?.Remove();

            footer = section.HeadersFooters[HeaderFooterType.FooterEven];
            footer?.Remove();
        }

        // Save the modified document as DOCX.
        doc.Save(outputPath);
    }
}
