using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class ContentControlValidator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Sample Form");

        // Insert a required plain‑text content control for "CustomerName".
        StructuredDocumentTag nameControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "required"
        };
        // Initially empty.
        nameControl.RemoveAllChildren();
        nameControl.AppendChild(new Run(doc, string.Empty));
        builder.InsertNode(nameControl);
        builder.Writeln(); // Move to next line.

        // Insert an optional plain‑text content control for "Comments".
        StructuredDocumentTag commentsControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Comments",
            Tag = "optional"
        };
        commentsControl.RemoveAllChildren();
        commentsControl.AppendChild(new Run(doc, string.Empty));
        builder.InsertNode(commentsControl);
        builder.Writeln();

        // Simulate user input: fill the required control, leave optional empty.
        nameControl.RemoveAllChildren();
        nameControl.AppendChild(new Run(doc, "Acme Corp"));

        // Validate required content controls before saving.
        var requiredControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                   .OfType<StructuredDocumentTag>()
                                   .Where(sdt => sdt.Tag == "required")
                                   .ToList();

        var validationErrors = requiredControls
            .Where(sdt => string.IsNullOrWhiteSpace(sdt.GetText()))
            .Select(sdt => new { sdt.Title, sdt.Tag })
            .ToList();

        if (validationErrors.Any())
        {
            // Serialize validation errors to JSON for debugging (optional).
            string json = JsonConvert.SerializeObject(validationErrors, Formatting.Indented);
            Console.WriteLine("Validation failed. Empty required content controls:");
            Console.WriteLine(json);
            throw new InvalidOperationException("One or more required content controls are empty.");
        }

        // All required controls contain text; save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ValidatedDocument.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
