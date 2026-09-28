using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class ContentControlPlaceholderReplacer
{
    public static void Main()
    {
        // Create a sample template document with placeholder content controls.
        Document template = new Document();
        Paragraph para = template.FirstSection.Body.FirstParagraph;

        // First placeholder content control.
        StructuredDocumentTag firstNameSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        firstNameSdt.Title = "FirstName";
        firstNameSdt.Tag = "first-name";
        firstNameSdt.RemoveAllChildren();
        firstNameSdt.AppendChild(new Run(template, "Enter first name"));
        para.AppendChild(firstNameSdt);

        // Add a space between placeholders.
        para.AppendChild(new Run(template, " "));

        // Second placeholder content control.
        StructuredDocumentTag lastNameSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        lastNameSdt.Title = "LastName";
        lastNameSdt.Tag = "last-name";
        lastNameSdt.RemoveAllChildren();
        lastNameSdt.AppendChild(new Run(template, "Enter last name"));
        para.AppendChild(lastNameSdt);

        // Save the template document.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Simulate user input values stored in a dictionary.
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" }
        };

        // Load the template document for processing.
        Document doc = new Document(templatePath);

        // Find all content controls (StructuredDocumentTag nodes) in the document.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Use the Title property as the key to look up replacement text.
            if (sdt.Title != null && userInputs.TryGetValue(sdt.Title, out string replacement))
            {
                // Replace the placeholder text with the user-provided value.
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, replacement));
            }
        }

        // Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
