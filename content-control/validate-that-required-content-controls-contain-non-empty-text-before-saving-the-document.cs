using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class ContentControlValidator
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Get the first paragraph of the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a required plain‑text content control (will be left empty).
        StructuredDocumentTag requiredSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        requiredSdt.Title = "RequiredName";
        requiredSdt.Tag = "required-name";
        requiredSdt.RemoveAllChildren(); // Ensure it has no content.
        paragraph.AppendChild(requiredSdt);

        // Create an optional plain‑text content control with some text.
        StructuredDocumentTag optionalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        optionalSdt.Title = "OptionalName";
        optionalSdt.Tag = "optional-name";
        optionalSdt.RemoveAllChildren();
        optionalSdt.AppendChild(new Run(doc, "Optional value"));
        paragraph.AppendChild(optionalSdt);

        // Attempt to validate and save the document.
        try
        {
            ValidateRequiredContentControls(doc);
            doc.Save("validated.docx");
            Console.WriteLine("Document saved successfully.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Validation failed: {ex.Message}");

            // Fill the required content control with a default value and retry.
            requiredSdt.RemoveAllChildren();
            requiredSdt.AppendChild(new Run(doc, "Default required value"));
            ValidateRequiredContentControls(doc);
            doc.Save("validated.docx");
            Console.WriteLine("Document saved after fixing required content controls.");
        }
    }

    private static void ValidateRequiredContentControls(Document doc)
    {
        // Find all content controls that are marked as required (by Title starting with "Required").
        var requiredControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.Title != null && sdt.Title.StartsWith("Required", StringComparison.Ordinal));

        foreach (var sdt in requiredControls)
        {
            // Get the visible text inside the content control.
            string text = sdt.GetText()?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(text))
            {
                throw new InvalidOperationException($"Content control titled '{sdt.Title}' must not be empty.");
            }
        }
    }
}
