using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with placeholder content controls.
        Document template = new Document();
        Paragraph para = template.FirstSection.Body.FirstParagraph;

        // Placeholder for "Name"
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        nameSdt.Title = "Name";
        nameSdt.Tag = "name";
        nameSdt.RemoveAllChildren();
        nameSdt.AppendChild(new Run(template, "{{Name}}"));
        para.AppendChild(nameSdt);

        // Add a space between placeholders.
        para.AppendChild(new Run(template, " "));

        // Placeholder for "Email"
        StructuredDocumentTag emailSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        emailSdt.Title = "Email";
        emailSdt.Tag = "email";
        emailSdt.RemoveAllChildren();
        emailSdt.AppendChild(new Run(template, "{{Email}}"));
        para.AppendChild(emailSdt);

        // Save the template.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Step 2: Load the template and replace placeholders with user input values.
        Document doc = new Document(templatePath);

        // Simulated user inputs.
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Name", "Alice Johnson" },
            { "Email", "alice.johnson@example.com" }
        };

        // Find all content controls in the document.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>();

        foreach (var entry in userInputs)
        {
            // Locate the content control by its Title (case‑insensitive).
            StructuredDocumentTag? targetSdt = sdtNodes.FirstOrDefault(s => string.Equals(s.Title, entry.Key, StringComparison.OrdinalIgnoreCase));

            if (targetSdt != null)
            {
                // Replace the placeholder text with the actual value.
                targetSdt.RemoveAllChildren();
                targetSdt.AppendChild(new Run(doc, entry.Value));
            }
        }

        // Save the result.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
