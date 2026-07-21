using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class ContentControlPlaceholderReplacement
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX file that contains plain‑text content
        // controls acting as placeholders.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // First placeholder – Title: "FirstName"
        StructuredDocumentTag firstNameTag = new StructuredDocumentTag(
            templateDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "FirstName",
            Tag = "first-name"
        };
        firstNameTag.RemoveAllChildren();
        firstNameTag.AppendChild(new Run(templateDoc, "<<FirstName>>"));
        builder.InsertNode(firstNameTag);
        builder.Writeln(); // move to next line

        // Second placeholder – Title: "LastName"
        StructuredDocumentTag lastNameTag = new StructuredDocumentTag(
            templateDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "LastName",
            Tag = "last-name"
        };
        lastNameTag.RemoveAllChildren();
        lastNameTag.AppendChild(new Run(templateDoc, "<<LastName>>"));
        builder.InsertNode(lastNameTag);
        builder.Writeln(); // move to next line

        // Third placeholder – Title: "Email"
        StructuredDocumentTag emailTag = new StructuredDocumentTag(
            templateDoc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Email",
            Tag = "email"
        };
        emailTag.RemoveAllChildren();
        emailTag.AppendChild(new Run(templateDoc, "<<Email>>"));
        builder.InsertNode(emailTag);
        builder.Writeln(); // finish paragraph

        // Save the template document to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Prepare a dictionary that maps placeholder titles to real values.
        // -----------------------------------------------------------------
        var userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName",  "Doe" },
            { "Email",     "john.doe@example.com" }
        };

        // (Optional) Serialize the dictionary to a JSON file for demonstration.
        File.WriteAllText("UserInputs.json", JsonConvert.SerializeObject(userInputs, Formatting.Indented));

        // -----------------------------------------------------------------
        // Step 3: Load the template document and replace each placeholder.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Enumerate all StructuredDocumentTag nodes in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Use the Title property as the lookup key; fall back to Tag if Title is missing.
            string key = sdt.Title ?? sdt.Tag;
            if (key != null && userInputs.TryGetValue(key, out string replacement))
            {
                // Clear existing children (placeholder text) and insert the new value.
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, replacement));
            }
        }

        // Save the resulting document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
    }
}
