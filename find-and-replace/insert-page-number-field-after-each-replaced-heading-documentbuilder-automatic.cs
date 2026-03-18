using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

class InsertPageNumbersAfterHeadings
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph uses a built‑in heading style.
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            bool isHeading = styleId == StyleIdentifier.Heading1 ||
                             styleId == StyleIdentifier.Heading2 ||
                             styleId == StyleIdentifier.Heading3 ||
                             styleId == StyleIdentifier.Heading4 ||
                             styleId == StyleIdentifier.Heading5 ||
                             styleId == StyleIdentifier.Heading6 ||
                             styleId == StyleIdentifier.Heading7 ||
                             styleId == StyleIdentifier.Heading8 ||
                             styleId == StyleIdentifier.Heading9;

            if (isHeading)
            {
                // Insert a PAGE field after the heading paragraph.
                // The overload inserts the field at the end of the paragraph (refNode = null) and
                // places it after the existing content (isAfter = true).
                para.InsertField(FieldType.FieldPage, true, null, true);
            }
        }

        // Update all fields so that the newly inserted PAGE fields show correct numbers.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
