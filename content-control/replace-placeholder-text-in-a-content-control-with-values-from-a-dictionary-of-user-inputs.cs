using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define placeholder keys that will be used as the Title of each content control.
        string[] placeholderKeys = { "Name", "Email", "Phone" };

        // Insert a paragraph and a plain‑text content control for each placeholder.
        foreach (string key in placeholderKeys)
        {
            // Write a label before the content control.
            builder.Writeln($"Please enter {key}:");

            // Create an inline plain‑text StructuredDocumentTag.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = key,                     // Use the key as the Title for lookup later.
                IsShowingPlaceholderText = true // Show placeholder text initially.
            };

            // Insert the content control into the document.
            builder.InsertNode(sdt);

            // Add placeholder text inside the content control.
            sdt.AppendChild(new Run(doc, $"[{key}]"));
        }

        // Save the intermediate document (optional, can be omitted).
        doc.Save("ContentControls_WithPlaceholders.docx");

        // Simulate user input values stored in a dictionary.
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Name", "John Doe" },
            { "Email", "john.doe@example.com" },
            { "Phone", "+1‑555‑123‑4567" }
        };

        // Replace placeholder text in each content control with the corresponding user input.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            if (sdt.Title != null && userInputs.TryGetValue(sdt.Title, out string replacement))
            {
                // Clear existing contents of the content control.
                sdt.RemoveAllChildren();

                // Insert the replacement text.
                sdt.AppendChild(new Run(doc, replacement));

                // Ensure the control no longer shows placeholder text.
                sdt.IsShowingPlaceholderText = false;
            }
        }

        // Save the final document with replaced values.
        doc.Save("ContentControls_Filled.docx");
    }
}
