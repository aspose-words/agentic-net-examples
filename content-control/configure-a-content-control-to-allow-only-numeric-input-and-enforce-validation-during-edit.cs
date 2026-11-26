using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure there is a paragraph to host the content control.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create an inline plain‑text content control.
        StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NumericInput",
            Tag = "numeric-input"
            // PlaceholderName omitted because it requires a building block that does not exist.
        };

        // Initialise with empty content.
        numericSdt.RemoveAllChildren();
        numericSdt.AppendChild(new Run(doc, string.Empty));

        // Insert the content control into the paragraph with surrounding text.
        paragraph.AppendChild(new Run(doc, "Please provide a number: "));
        paragraph.AppendChild(numericSdt);
        paragraph.AppendChild(new Run(doc, "."));

        // Save the initial document.
        doc.Save("NumericContentControl.docx");

        // ----- Simulate editing: valid numeric input -----
        SetContentControlText(numericSdt, "12345");
        ValidateNumericContentControl(numericSdt, "Valid numeric input");

        // ----- Simulate editing: invalid non‑numeric input -----
        SetContentControlText(numericSdt, "ABC");
        try
        {
            ValidateNumericContentControl(numericSdt, "Invalid non‑numeric input");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Validation error: {ex.Message}");
        }

        // Save the document after simulated edits.
        doc.Save("NumericContentControl_Edited.docx");
    }

    private static void SetContentControlText(StructuredDocumentTag sdt, string text)
    {
        // Replace any existing children with a single run containing the new text.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(sdt.Document, text));
    }

    private static void ValidateNumericContentControl(StructuredDocumentTag sdt, string context)
    {
        string content = sdt.GetText().Trim();
        if (!Regex.IsMatch(content, @"^\d+$"))
        {
            throw new InvalidOperationException($"{context}: Content '{content}' is not numeric.");
        }

        Console.WriteLine($"{context}: Content '{content}' is valid numeric.");
    }
}
