using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a required plain‑text content control.
        StructuredDocumentTag requiredSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "RequiredName",   // Title indicates that this control is required.
            Tag = "Name"
        };
        // Add some initial text so the validation will succeed.
        requiredSdt.AppendChild(new Run(doc, "John Doe"));
        builder.InsertNode(requiredSdt);

        // Insert an optional plain‑text content control.
        StructuredDocumentTag optionalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "OptionalComment",
            Tag = "Comment"
        };
        // Leave this control empty to demonstrate that optional controls are ignored.
        builder.InsertNode(optionalSdt);

        // Validate required content controls before saving.
        ValidateRequiredContentControls(doc);

        // Save the document.
        doc.Save("ValidatedDocument.docx");
    }

    private static void ValidateRequiredContentControls(Document doc)
    {
        // Retrieve all StructuredDocumentTag nodes in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        foreach (Node node in sdtNodes)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)node;

            // Consider a control "required" if its Title starts with "Required".
            if (sdt.Title != null && sdt.Title.StartsWith("Required", StringComparison.OrdinalIgnoreCase))
            {
                // Get the visible text inside the content control.
                string text = sdt.GetText()?.Trim() ?? string.Empty;

                // If the text is empty, throw an exception.
                if (string.IsNullOrEmpty(text))
                {
                    throw new InvalidOperationException(
                        $"The required content control '{sdt.Title}' (Tag: {sdt.Tag}) must contain non‑empty text.");
                }
            }
        }
    }
}
