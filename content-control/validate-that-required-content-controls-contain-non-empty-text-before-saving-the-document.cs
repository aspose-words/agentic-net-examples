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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an introductory paragraph.
        builder.Writeln("Please fill in the following form:");

        // Insert a plain‑text content control for "Name" (required) and pre‑fill it.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Name",
            Tag = "required"
        };
        nameSdt.RemoveAllChildren();
        nameSdt.AppendChild(new Run(doc, "John Doe"));
        builder.InsertNode(nameSdt);

        // Add a space between controls.
        builder.Write(" ");

        // Insert a plain‑text content control for "Email" (required) – left empty.
        StructuredDocumentTag emailSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Email",
            Tag = "required"
        };
        emailSdt.RemoveAllChildren(); // No initial text.
        builder.InsertNode(emailSdt);

        // Validation: ensure every required content control contains non‑empty text.
        var validationResults = new List<ValidationResult>();
        bool anyInvalid = false;

        // Enumerate all StructuredDocumentTag nodes in the document.
        var allSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                        .Cast<StructuredDocumentTag>();

        foreach (StructuredDocumentTag sdt in allSdt)
        {
            // Consider a control required if its Tag is set to "required".
            if (string.Equals(sdt.Tag, "required", StringComparison.OrdinalIgnoreCase))
            {
                string text = sdt.GetText().Trim();
                bool isValid = !string.IsNullOrEmpty(text);
                if (!isValid) anyInvalid = true;

                validationResults.Add(new ValidationResult
                {
                    Title = sdt.Title,
                    IsValid = isValid,
                    Message = isValid ? "OK" : "Content is empty."
                });
            }
        }

        // Serialize validation results to a JSON file.
        string json = JsonConvert.SerializeObject(validationResults, Formatting.Indented);
        File.WriteAllText("validation.json", json);

        // If any required control is empty, report and abort saving without throwing an unhandled exception.
        if (anyInvalid)
        {
            Console.WriteLine("One or more required content controls are empty. See validation.json for details.");
            return; // Exit gracefully.
        }

        // All required controls are valid – save the document.
        doc.Save("validated.docx");
        Console.WriteLine("Document saved successfully as validated.docx");
    }

    private class ValidationResult
    {
        public string Title { get; set; } = string.Empty;
        public bool IsValid { get; set; }
        public string Message { get; set; } = string.Empty;
    }
}
