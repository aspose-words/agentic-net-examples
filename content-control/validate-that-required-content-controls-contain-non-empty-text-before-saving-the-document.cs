using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph of the document body.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // ----- Required content control with non‑empty text -----
        StructuredDocumentTag requiredSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        requiredSdt.Title = "CustomerName";
        requiredSdt.Tag = "required";
        requiredSdt.RemoveAllChildren();
        requiredSdt.AppendChild(new Run(doc, "Contoso Ltd."));
        firstParagraph.AppendChild(requiredSdt);

        // ----- Optional content control (can be empty) -----
        StructuredDocumentTag optionalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        optionalSdt.Title = "OptionalNote";
        optionalSdt.Tag = "optional";
        optionalSdt.RemoveAllChildren();
        // No text added – this control is intentionally left empty.
        firstParagraph.AppendChild(optionalSdt);

        // ----- Required content control that is empty (to demonstrate validation failure) -----
        StructuredDocumentTag emptyRequiredSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        emptyRequiredSdt.Title = "EmptyRequired";
        emptyRequiredSdt.Tag = "required";
        emptyRequiredSdt.RemoveAllChildren();
        // No text added – this will cause validation to fail.
        firstParagraph.AppendChild(emptyRequiredSdt);

        // Validate required content controls before saving.
        try
        {
            ValidateRequiredContentControls(doc);
            // If validation passes, save the document.
            string outputPath = "validated.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        catch (InvalidOperationException ex)
        {
            // Validation failed – report the issue.
            Console.WriteLine($"Validation error: {ex.Message}");
        }
    }

    private static void ValidateRequiredContentControls(Document document)
    {
        // Find all StructuredDocumentTag nodes that are marked as required via the Tag property.
        var requiredControls = document.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => string.Equals(sdt.Tag, "required", StringComparison.OrdinalIgnoreCase));

        foreach (var sdt in requiredControls)
        {
            // Get the visible text inside the content control.
            string text = sdt.GetText()?.Trim() ?? string.Empty;

            // If the text is empty, throw an exception indicating which control failed.
            if (string.IsNullOrEmpty(text))
            {
                string title = string.IsNullOrEmpty(sdt.Title) ? "(no title)" : sdt.Title;
                throw new InvalidOperationException($"Content control '{title}' (Tag='{sdt.Tag}') is required but contains no text.");
            }
        }
    }
}
