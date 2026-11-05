using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with placeholder content controls.
        Document template = new Document();
        Paragraph templateParagraph = template.FirstSection.Body.FirstParagraph;

        // Placeholder for First Name
        StructuredDocumentTag firstNameSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        firstNameSdt.Title = "FirstName";
        firstNameSdt.Tag = "first-name";
        firstNameSdt.RemoveAllChildren();
        firstNameSdt.AppendChild(new Run(template, "<<FirstName>>"));
        templateParagraph.AppendChild(firstNameSdt);

        // Placeholder for Last Name
        StructuredDocumentTag lastNameSdt = new StructuredDocumentTag(template, SdtType.PlainText, MarkupLevel.Inline);
        lastNameSdt.Title = "LastName";
        lastNameSdt.Tag = "last-name";
        lastNameSdt.RemoveAllChildren();
        lastNameSdt.AppendChild(new Run(template, "<<LastName>>"));
        templateParagraph.AppendChild(lastNameSdt);

        // Save the template document.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Step 2: Define user input values.
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" }
        };

        // Step 3: Load the template and replace placeholders.
        Document doc = new Document(templatePath);
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                          .OfType<StructuredDocumentTag>();

        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            if (sdt.Title != null && userInputs.TryGetValue(sdt.Title, out string replacement))
            {
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, replacement));
            }
        }

        // Step 4: Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: Write a JSON file summarizing the replacements.
        var summary = sdtNodes
            .Where(s => s.Title != null && userInputs.ContainsKey(s.Title))
            .Select(s => new { Title = s.Title, ReplacedWith = userInputs[s.Title] })
            .ToList();

        string json = JsonConvert.SerializeObject(summary, Formatting.Indented);
        File.WriteAllText("replacements.json", json);
    }
}
