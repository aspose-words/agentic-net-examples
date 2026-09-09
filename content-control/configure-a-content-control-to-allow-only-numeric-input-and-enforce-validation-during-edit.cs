using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Fields;

namespace ContentControlNumericExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the content control.
            builder.Writeln("Please enter a numeric value:");

            // Create a block‑level plain‑text content control.
            StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
            {
                Title = "NumericOnly",
                Tag = "numeric"
            };

            // Add a paragraph inside the content control – this will host the numeric field.
            Paragraph sdtParagraph = new Paragraph(doc);
            numericSdt.AppendChild(sdtParagraph);

            // Insert the content control into the document body.
            doc.FirstSection.Body.AppendChild(numericSdt);

            // Move the builder cursor to the paragraph inside the content control.
            builder.MoveTo(sdtParagraph);

            // Insert a text input form field that only accepts numbers.
            // The field type TextFormFieldType.Number enforces numeric input during editing.
            builder.InsertTextInput("NumberField", TextFormFieldType.Number, "", "0", 10);

            // Prevent the user from deleting the content control itself.
            numericSdt.LockContentControl = true;

            // Save the resulting document.
            doc.Save("NumericContentControl.docx");
        }
    }
}
