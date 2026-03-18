using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;

class ContentControlHyperlinkDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Create a plain‑text content control (StructuredDocumentTag).
        // The DocumentBuilder.StartStructuredDocumentTag API is only
        // available in newer Aspose.Words versions. For older versions
        // we create the tag manually and insert it into the document.
        // ------------------------------------------------------------
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        builder.InsertNode(sdt);
        // Move the cursor inside the newly created content control.
        builder.MoveTo(sdt);

        // Insert a hyperlink field inside the content control.
        // Display text: "Visit Aspose"
        // Target URL:  "https://www.aspose.com"
        // The third argument (false) indicates that the second parameter is a URL, not a bookmark.
        Field hyperlinkField = builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // Cast to FieldHyperlink to set additional properties if needed.
        if (hyperlinkField is FieldHyperlink fieldHyperlink)
        {
            // Ensure the link opens in a new window (optional).
            fieldHyperlink.OpenInNewWindow = true;
        }

        // Move the cursor out of the content control so that subsequent operations
        // (e.g., saving) are not performed inside the tag.
        builder.MoveToDocumentEnd();

        // Update fields so that the hyperlink result is calculated.
        doc.UpdateFields();

        // Save the document in DOCX format.
        const string docxPath = "ContentControlHyperlink.docx";
        doc.Save(docxPath);

        // Convert the document to PDF.
        const string pdfPath = "ContentControlHyperlink.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Reload the DOCX to verify that the hyperlink target is still correct.
        Document loadedDoc = new Document(docxPath);
        // Locate the first FieldHyperlink inside any StructuredDocumentTag.
        FieldHyperlink loadedHyperlink = null;
        foreach (StructuredDocumentTag tag in loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true))
        {
            foreach (Field field in tag.Range.Fields)
            {
                if (field.Type == FieldType.FieldHyperlink)
                {
                    loadedHyperlink = (FieldHyperlink)field;
                    break;
                }
            }
            if (loadedHyperlink != null) break;
        }

        // Verify that the hyperlink address matches the expected URL.
        const string expectedUrl = "https://www.aspose.com";
        if (loadedHyperlink != null && loadedHyperlink.Address == expectedUrl)
        {
            Console.WriteLine("Verification succeeded: hyperlink address is correct after conversion.");
        }
        else
        {
            Console.WriteLine("Verification failed: hyperlink address is missing or incorrect.");
        }
    }
}
