using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace ContentControlValidationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();

            // Get the first paragraph of the document body.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

            // Create a required plain‑text content control (initially empty).
            StructuredDocumentTag requiredSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
            requiredSdt.Title = "CustomerName";
            requiredSdt.Tag = "required";
            requiredSdt.RemoveAllChildren(); // Ensure it is empty.
            paragraph.AppendChild(requiredSdt);

            // Add a space between controls.
            paragraph.AppendChild(new Run(doc, " "));

            // Create an optional plain‑text content control with some text.
            StructuredDocumentTag optionalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
            optionalSdt.Title = "OptionalInfo";
            optionalSdt.Tag = "optional";
            optionalSdt.RemoveAllChildren();
            optionalSdt.AppendChild(new Run(doc, "Sample data"));
            paragraph.AppendChild(optionalSdt);

            // Attempt to validate and save the document.
            try
            {
                ValidateRequiredContentControls(doc);
                doc.Save("validated.docx");
                Console.WriteLine("Document saved successfully.");
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine($"Validation failed: {ex.Message}");

                // Fix the required content control by adding non‑empty text.
                requiredSdt.RemoveAllChildren();
                requiredSdt.AppendChild(new Run(doc, "John Doe"));

                // Re‑validate and save.
                ValidateRequiredContentControls(doc);
                doc.Save("validated.docx");
                Console.WriteLine("Document saved after fixing required content control.");
            }
        }

        private static void ValidateRequiredContentControls(Document doc)
        {
            var requiredControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(sdt => sdt.Tag == "required");

            foreach (var sdt in requiredControls)
            {
                string text = sdt.GetText().Trim();
                if (string.IsNullOrEmpty(text))
                {
                    throw new InvalidOperationException(
                        $"Content control '{sdt.Title}' is required but contains no text.");
                }
            }
        }
    }
}
