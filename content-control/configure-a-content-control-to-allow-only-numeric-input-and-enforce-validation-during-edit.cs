using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Fields;

namespace NumericContentControlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an introductory paragraph.
            builder.Writeln("Please enter a numeric value:");

            // Create an inline plain‑text content control.
            StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "NumericInput",
                Tag = "numeric-input",
                // Prevent the user from deleting the content control, but allow editing its contents.
                LockContentControl = true
            };

            // Insert the content control into the current paragraph.
            builder.InsertNode(numericSdt);

            // Move the builder's cursor inside the newly inserted content control.
            builder.MoveTo(numericSdt);

            // Insert a text input form field that accepts only numbers.
            // Parameters: name, field type, default text, placeholder text, max length.
            builder.InsertTextInput("NumericField", TextFormFieldType.Number, "", "0", 10);

            // Save the document to the working directory.
            const string outputPath = "NumericContentControl.docx";
            doc.Save(outputPath);

            // Optional: Load the document again and verify that the field is of numeric type.
            Document loadedDoc = new Document(outputPath);
            FormField numericField = loadedDoc.Range.FormFields["NumericField"];
            if (numericField == null || numericField.Type != FieldType.FieldFormTextInput)
            {
                throw new InvalidOperationException("Numeric form field was not created correctly.");
            }

            // The example finishes without requiring any user interaction.
        }
    }
}
