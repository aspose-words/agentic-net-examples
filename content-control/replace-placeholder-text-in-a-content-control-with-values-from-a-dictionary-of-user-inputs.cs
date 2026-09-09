using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace ContentControlPlaceholderReplacement
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading.
            builder.Writeln("Invoice");

            // Insert a plain‑text content control for the customer name.
            StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "customer-name"
            };
            nameSdt.RemoveAllChildren();
            nameSdt.AppendChild(new Run(doc, "Enter name"));
            builder.Write("Customer: ");
            builder.InsertNode(nameSdt);
            builder.Writeln();

            // Insert a plain‑text content control for the address.
            StructuredDocumentTag addressSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "Address",
                Tag = "address"
            };
            addressSdt.RemoveAllChildren();
            addressSdt.AppendChild(new Run(doc, "Enter address"));
            builder.Write("Address: ");
            builder.InsertNode(addressSdt);
            builder.Writeln();

            // Save the seed document (optional, shows the original placeholders).
            doc.Save("Invoice_Seed.docx");

            // Dictionary that simulates user input values.
            Dictionary<string, string> userInputs = new Dictionary<string, string>
            {
                { "CustomerName", "John Doe" },
                { "Address", "123 Main St, Springfield" }
            };

            // Replace placeholder text in each content control whose Title matches a key in the dictionary.
            foreach (StructuredDocumentTag sdt in doc.GetChildNodes(NodeType.StructuredDocumentTag, true).OfType<StructuredDocumentTag>())
            {
                if (sdt.Title != null && userInputs.TryGetValue(sdt.Title, out string replacement))
                {
                    sdt.RemoveAllChildren();
                    sdt.AppendChild(new Run(doc, replacement));
                    sdt.IsShowingPlaceholderText = false; // Ensure the control shows the actual text.
                }
            }

            // Save the updated document.
            doc.Save("Invoice_Updated.docx");

            // (Optional) Serialize the user inputs to a JSON file for reference.
            string json = JsonConvert.SerializeObject(userInputs, Formatting.Indented);
            File.WriteAllText("UserInputs.json", json);
        }
    }
}
