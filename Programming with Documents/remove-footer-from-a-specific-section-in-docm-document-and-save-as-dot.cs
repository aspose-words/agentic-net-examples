using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCM document.
        Document doc = new Document("Input.docm");

        // Index of the section whose footer should be removed (0‑based).
        int targetSectionIndex = 1; // change as required

        if (targetSectionIndex >= 0 && targetSectionIndex < doc.Sections.Count)
        {
            Section section = doc.Sections[targetSectionIndex];

            // Remove the primary footer, if it exists.
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            footer?.Remove();

            // Remove the first‑page footer, if it exists.
            footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
            footer?.Remove();

            // Remove the even‑page footer, if it exists.
            footer = section.HeadersFooters[HeaderFooterType.FooterEven];
            footer?.Remove();
        }

        // Save the modified document as a DOTM (macro‑enabled template).
        doc.Save("Output.dotm", SaveFormat.Dotm);
    }
}
