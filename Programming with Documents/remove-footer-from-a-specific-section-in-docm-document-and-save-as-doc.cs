using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveFooterFromSection
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\source.docm";

        // Path where the resulting DOCX will be saved.
        string outputPath = @"C:\Docs\result.docx";

        // Zero‑based index of the section whose footer should be removed.
        int targetSectionIndex = 1; // change as needed

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Ensure the requested section exists.
        if (targetSectionIndex >= 0 && targetSectionIndex < doc.Sections.Count)
        {
            Section section = doc.Sections[targetSectionIndex];

            // Option 1: Clear the contents of all headers and footers in this section.
            // This keeps the HeaderFooter objects but makes them empty, causing the
            // section to inherit footers from the previous section.
            section.ClearHeadersFooters();

            // Option 2: Completely remove footer objects from this section.
            // Remove primary footer.
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            footer?.Remove();

            // Remove first‑page footer.
            footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
            footer?.Remove();

            // Remove even‑page footer.
            footer = section.HeadersFooters[HeaderFooterType.FooterEven];
            footer?.Remove();
        }

        // Save the modified document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
