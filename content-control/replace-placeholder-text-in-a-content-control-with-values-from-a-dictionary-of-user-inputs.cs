using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Customer Information:");

        // Insert a plain‑text content control for the customer's name.
        StructuredDocumentTag nameTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        nameTag.RemoveAllChildren();
        nameTag.AppendChild(new Run(doc, "Enter name"));
        builder.InsertNode(nameTag);
        builder.Writeln(); // Move to the next line.

        // Insert a plain‑text content control for the customer's address.
        StructuredDocumentTag addressTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Address",
            Tag = "address"
        };
        addressTag.RemoveAllChildren();
        addressTag.AppendChild(new Run(doc, "Enter address"));
        builder.InsertNode(addressTag);
        builder.Writeln();

        // Dictionary that simulates user input values.
        Dictionary<string, string> userInputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "CustomerName", "John Doe" },
            { "Address", "123 Main St, Anytown" }
        };

        // Find all content controls in the document and replace their placeholder text
        // with the corresponding values from the dictionary.
        foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                                .OfType<StructuredDocumentTag>())
        {
            // Use the Title property as the lookup key.
            if (userInputs.TryGetValue(sdt.Title, out string replacement))
            {
                sdt.RemoveAllChildren();
                sdt.AppendChild(new Run(doc, replacement));
            }
        }

        // Save the resulting document.
        doc.Save("output.docx");
    }
}
