using System;
using Aspose.Words;
using Aspose.Words.Tables;

class HeaderCopyExample
{
    static void Main()
    {
        // Load an existing document (replace with your source file path)
        Document doc = new Document("Input.docx");

        // Iterate through all sections starting from the second one
        for (int i = 1; i < doc.Sections.Count; i++)
        {
            Section previousSection = doc.Sections[i - 1];
            Section currentSection = doc.Sections[i];

            // Get the primary header from the previous section
            HeaderFooter previousHeader = previousSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (previousHeader != null)
            {
                // Clone the header (deep clone) so it can be added to the new section
                HeaderFooter clonedHeader = (HeaderFooter)previousHeader.Clone(true);

                // Add the cloned header to the current section's HeadersFooters collection
                // This copies the header content without linking it to the previous section
                currentSection.HeadersFooters.Add(clonedHeader);
            }
        }

        // Save the modified document (replace with your desired output file path)
        doc.Save("Output.docx");
    }
}
