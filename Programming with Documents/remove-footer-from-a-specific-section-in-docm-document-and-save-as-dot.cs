using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Load the source DOCM document.
        Document doc = new Document("Input.docm");

        // Index of the section whose footer should be removed (0‑based).
        int sectionIndex = 1; // change as needed

        if (sectionIndex >= 0 && sectionIndex < doc.Sections.Count)
        {
            Section targetSection = doc.Sections[sectionIndex];

            // Remove the primary footer if present.
            HeaderFooter footer = targetSection.HeadersFooters[HeaderFooterType.FooterPrimary];
            footer?.Remove();

            // Remove the first‑page footer if present.
            footer = targetSection.HeadersFooters[HeaderFooterType.FooterFirst];
            footer?.Remove();

            // Remove the even‑page footer if present.
            footer = targetSection.HeadersFooters[HeaderFooterType.FooterEven];
            footer?.Remove();
        }

        // Save the modified document as a DOTM (macro‑enabled template).
        doc.Save("Output.dotm");
    }
}
