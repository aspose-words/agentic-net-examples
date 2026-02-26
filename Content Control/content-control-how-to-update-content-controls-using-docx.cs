using System;
using Aspose.Words;
using Aspose.Words.Markup;

class UpdateContentControls
{
    static void Main()
    {
        // Load the existing DOCX file that contains content controls (structured document tags).
        Document doc = new Document("InputWithContentControls.docx");

        // Retrieve all content controls in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control and update its contents as needed.
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Example: Update a plain‑text content control with the title "CustomerName".
            if (sdt.Title == "CustomerName")
            {
                // Remove any existing child nodes (e.g., previous text).
                sdt.RemoveAllChildren();

                // Insert the new text into the content control.
                sdt.AppendChild(new Run(doc, "John Doe"));
            }

            // Example: Update a plain‑text content control with the tag "Address".
            if (sdt.Tag == "Address")
            {
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, "123 Main Street, Springfield"));
            }

            // Example: For a dropdown list content control, set the selected value.
            if (sdt.SdtType == SdtType.DropDownList && sdt.Title == "Country")
            {
                // The list items are stored in the SdtListItemCollection.
                // Choose the item with the desired display text (e.g., "USA").
                foreach (SdtListItem item in sdt.ListItems)
                {
                    if (item.DisplayText == "USA")
                    {
                        // The selected value is stored in the SdtListItem's Value property.
                        sdt.RemoveAllChildren();
                        sdt.AppendChild(new Run(doc, item.Value));
                        break;
                    }
                }
            }
        }

        // Save the modified document to a new file.
        doc.Save("UpdatedContentControls.docx");
    }
}
