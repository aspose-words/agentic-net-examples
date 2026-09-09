using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Prepare a DocumentBuilder for inserting nodes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Required content control #1 -----
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,
            MarkupLevel.Inline);
        nameSdt.Title = "CustomerName";
        nameSdt.Tag = "customer-name";
        nameSdt.RemoveAllChildren();
        nameSdt.AppendChild(new Run(doc, "John Doe"));
        builder.InsertNode(nameSdt);

        // ----- Required content control #2 (intentionally left empty) -----
        StructuredDocumentTag emailSdt = new StructuredDocumentTag(
            doc,
            SdtType.PlainText,
            MarkupLevel.Inline);
        emailSdt.Title = "Email";
        emailSdt.Tag = "email";
        emailSdt.RemoveAllChildren();
        // No text added here to simulate an empty required field.
        builder.InsertNode(emailSdt);

        // Validate that all required content controls contain non‑empty text.
        // If validation fails, the document will not be saved.
        try
        {
            ValidateRequiredContentControls(doc, new[] { "CustomerName", "Email" });
            // Save the validated document only when validation succeeds.
            doc.Save("validated.docx");
            Console.WriteLine("Document saved successfully.");
        }
        catch (InvalidOperationException ex)
        {
            // Report validation error without crashing the program.
            Console.WriteLine($"Validation error: {ex.Message}");
        }
    }

    private static void ValidateRequiredContentControls(Document doc, string[] requiredTitles)
    {
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (Node node in sdtNodes)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)node;
            // Check only the content controls whose Title is in the required list.
            if (Array.Exists(requiredTitles, title => title == sdt.Title))
            {
                string text = sdt.GetText().Trim();
                if (string.IsNullOrEmpty(text))
                {
                    throw new InvalidOperationException(
                        $"Content control '{sdt.Title}' must not be empty.");
                }
            }
        }
    }
}
