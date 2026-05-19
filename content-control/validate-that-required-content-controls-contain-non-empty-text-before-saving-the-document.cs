using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class ValidateContentControls
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a required plain‑text content control.
        StructuredDocumentTag requiredSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "RequiredField",
            Tag = "required"
        };
        // The control is left empty to demonstrate validation.
        builder.InsertNode(requiredSdt);
        builder.Writeln(); // Move to a new paragraph.

        // Insert an optional plain‑text content control and give it a value.
        StructuredDocumentTag optionalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "OptionalField",
            Tag = "optional"
        };
        optionalSdt.RemoveAllChildren();
        optionalSdt.AppendChild(new Run(doc, "Sample value"));
        builder.InsertNode(optionalSdt);
        builder.Writeln();

        // Validate that all required content controls contain non‑empty text.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>();

        foreach (var sdt in sdtNodes)
        {
            // Identify required controls by Tag or Title.
            bool isRequired = sdt.Tag == "required" || sdt.Title == "RequiredField";

            if (isRequired)
            {
                // Get the visible text inside the control.
                string text = sdt.GetText().Trim();

                if (string.IsNullOrEmpty(text))
                {
                    throw new InvalidOperationException(
                        $"The required content control '{sdt.Title}' (Tag='{sdt.Tag}') is empty.");
                }
            }
        }

        // All validations passed – save the document.
        doc.Save("ValidatedDocument.docx");
    }
}
