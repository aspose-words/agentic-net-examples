using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create an inline plain‑text content control.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

            // Set the friendly name (Title) and the non‑visible identifier (Tag).
            sdt.Title = "CustomerName";
            sdt.Tag = "customer-name";

            // Add placeholder text inside the control.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "Enter name here"));

            // Insert the content control into the first paragraph of the document.
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(sdt);

            // Save the document with the configured content control.
            doc.Save("ContentControlTitleTag.docx");

            // ---------- Later identification ----------
            // Locate the control by its Title.
            IStructuredDocumentTag? foundByTitle = doc.Range.StructuredDocumentTags.GetByTitle("CustomerName");
            if (foundByTitle != null && foundByTitle is StructuredDocumentTag titleSdt)
            {
                // Replace the placeholder text with actual content.
                titleSdt.RemoveAllChildren();
                titleSdt.AppendChild(new Run(doc, "Contoso Ltd."));
            }

            // Locate the control by its Tag.
            IStructuredDocumentTag? foundByTag = doc.Range.StructuredDocumentTags.GetByTag("customer-name");
            if (foundByTag != null && foundByTag is StructuredDocumentTag tagSdt)
            {
                // Change the appearance to show tags around the control.
                tagSdt.Appearance = SdtAppearance.Tags;
            }

            // Save the modified document.
            doc.Save("ContentControlTitleTag_Modified.docx");
        }
    }
}
