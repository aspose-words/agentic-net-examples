using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with plain‑text content controls that contain placeholder text.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a heading.
        builder.Writeln("User Information:");

        // ---- Name content control ----
        // Create a new paragraph for the control.
        Paragraph nameParagraph = template.FirstSection.Body.LastParagraph;
        // Create an inline plain‑text SDT.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Name",
            Tag = "name"
        };
        // Set placeholder text.
        nameSdt.RemoveAllChildren();
        nameSdt.AppendChild(new Run(template, "Enter name"));
        // Insert the SDT into the paragraph.
        nameParagraph.AppendChild(nameSdt);
        // Add a line break after the control.
        builder.Writeln();

        // ---- Email content control ----
        Paragraph emailParagraph = template.FirstSection.Body.LastParagraph;
        StructuredDocumentTag emailSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Email",
            Tag = "email"
        };
        emailSdt.RemoveAllChildren();
        emailSdt.AppendChild(new Run(template, "Enter email"));
        emailParagraph.AppendChild(emailSdt);
        builder.Writeln();

        // Save the template document.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Step 2: Define user input values in a dictionary.
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "name", "John Doe" },
            { "email", "john.doe@example.com" }
        };

        // Step 3: Load the template and replace placeholder text in each content control.
        Document doc = new Document(templatePath);

        // Find all StructuredDocumentTag nodes in the document.
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes.OfType<StructuredDocumentTag>())
        {
            // Use the Tag property as the lookup key; fall back to Title if Tag is empty.
            string key = !string.IsNullOrEmpty(sdt.Tag) ? sdt.Tag : sdt.Title;
            if (key != null && userInputs.TryGetValue(key, out string replacement))
            {
                // Replace the existing placeholder/run with the user‑provided value.
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, replacement));
            }
        }

        // Step 4: Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
