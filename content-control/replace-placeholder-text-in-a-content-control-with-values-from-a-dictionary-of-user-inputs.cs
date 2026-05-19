using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with plain‑text content controls acting as placeholders.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // First placeholder – Title: "FirstName"
        StructuredDocumentTag firstNameTag = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstName",
            Tag = "first-name"
        };
        firstNameTag.RemoveAllChildren();
        firstNameTag.AppendChild(new Run(template, "Enter first name"));
        builder.Writeln("Dear ");
        builder.InsertNode(firstNameTag);
        builder.Writeln(",");

        // Second placeholder – Title: "LastName"
        StructuredDocumentTag lastNameTag = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "LastName",
            Tag = "last-name"
        };
        lastNameTag.RemoveAllChildren();
        lastNameTag.AppendChild(new Run(template, "Enter last name"));
        builder.Writeln("Your surname is ");
        builder.InsertNode(lastNameTag);
        builder.Writeln(".");

        // Save the template for demonstration purposes.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Step 2: Load the document that contains the placeholders.
        Document doc = new Document(templatePath);

        // Step 3: Define user input values that should replace the placeholders.
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" }
        };

        // Optional: serialize the input dictionary to JSON (demonstrates the required package).
        string json = JsonConvert.SerializeObject(userInputs, Formatting.Indented);
        File.WriteAllText("userInputs.json", json);

        // Step 4: Replace each content control's text with the corresponding value from the dictionary.
        foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true).OfType<StructuredDocumentTag>())
        {
            if (sdt.Title != null && userInputs.TryGetValue(sdt.Title, out string replacement))
            {
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, replacement));
                // Ensure the placeholder is not shown after replacement.
                sdt.IsShowingPlaceholderText = false;
            }
        }

        // Step 5: Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
