using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a plain‑text content control that is intended for numeric input.
        StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NumericInput",
            Tag = "numeric-input"
            // PlaceholderName is omitted because a placeholder building block is not defined.
        };

        // Set an initial valid numeric value.
        numericSdt.RemoveAllChildren();
        numericSdt.AppendChild(new Run(doc, "123"));

        // Add the content control to the first paragraph of the document.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(numericSdt);

        // Save the intermediate document.
        string intermediatePath = "numeric_input_initial.docx";
        doc.Save(intermediatePath);

        // -----------------------------------------------------------------
        // Simulate user editing: load the document and change the content
        // control's text to an invalid (non‑numeric) value.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(intermediatePath);

        // Find the content control by its title.
        StructuredDocumentTag? foundSdt = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .FirstOrDefault(s => s.Title == "NumericInput");

        if (foundSdt == null)
        {
            throw new InvalidOperationException("Numeric input content control not found.");
        }

        // Replace the existing text with an invalid value to simulate editing.
        foundSdt.RemoveAllChildren();
        foundSdt.AppendChild(new Run(loadedDoc, "ABC")); // Non‑numeric input.

        // Validate that the content control contains only digits.
        string sdtText = foundSdt.GetText().Trim();

        if (!Regex.IsMatch(sdtText, @"^\d+$"))
        {
            // If validation fails, replace with a default numeric value and lock the control.
            foundSdt.RemoveAllChildren();
            foundSdt.AppendChild(new Run(loadedDoc, "0"));
            foundSdt.LockContents = true;
        }

        // Save the final document after validation.
        string finalPath = "numeric_input_validated.docx";
        loadedDoc.Save(finalPath);
    }
}
