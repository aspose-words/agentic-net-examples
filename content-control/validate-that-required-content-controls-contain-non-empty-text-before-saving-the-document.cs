using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlValidationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some introductory text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Please fill the form below:");

            // ---- Insert a required plain‑text content control with a non‑empty value ----
            StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "Name",
                Tag = "required"
            };
            nameSdt.RemoveAllChildren();                     // Ensure the control is empty before adding content.
            nameSdt.AppendChild(new Run(doc, "John Doe"));   // Sample non‑empty text.

            // ---- Insert a required plain‑text content control that is also filled ----
            StructuredDocumentTag emailSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "Email",
                Tag = "required"
            };
            emailSdt.RemoveAllChildren();                    // Ensure the control is empty before adding content.
            emailSdt.AppendChild(new Run(doc, "john.doe@example.com")); // Provide a non‑empty value.

            // Add the content controls to the document.
            builder.InsertNode(nameSdt);
            builder.Write(" "); // Separator.
            builder.InsertNode(emailSdt);
            builder.Writeln(); // End of line.

            // -------------------- Validation before saving --------------------
            // Find all StructuredDocumentTag nodes that are marked as required (Tag == "required").
            var requiredControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                                      .Cast<StructuredDocumentTag>()
                                      .Where(sdt => string.Equals(sdt.Tag, "required", StringComparison.OrdinalIgnoreCase));

            foreach (var sdt in requiredControls)
            {
                // Get the visible text inside the content control.
                string text = sdt.GetText().Trim();

                // If the text is empty, throw an exception to indicate validation failure.
                if (string.IsNullOrEmpty(text))
                {
                    throw new InvalidOperationException(
                        $"The required content control '{sdt.Title}' is empty. Please provide a value before saving.");
                }
            }

            // -------------------- Save the validated document --------------------
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "output");
            Directory.CreateDirectory(outputDir);

            string outputPath = Path.Combine(outputDir, "validated.docx");
            doc.Save(outputPath);
        }
    }
}
