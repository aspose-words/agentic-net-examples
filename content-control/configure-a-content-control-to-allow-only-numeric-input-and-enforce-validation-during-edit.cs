using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph that will contain the content control.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please enter a numeric value:");

        // Create an inline plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NumericInput",
            Tag = "numeric-input"
        };

        // Append the content control to the first paragraph.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(numericSdt);

        // Move the builder's cursor inside the newly created content control.
        builder.MoveTo(numericSdt);

        // Insert a text input form field that only accepts numbers.
        // The InsertTextInput overload does not have a 'placeholderText' parameter.
        // The fourth argument is the default field value that appears when the field is empty.
        builder.InsertTextInput(
            name: "NumericField",
            type: TextFormFieldType.Number,
            format: "",
            fieldValue: "0",   // default displayed value (acts as a placeholder)
            maxLength: 10);    // limit the number of characters the user can type

        // Save the document.
        doc.Save("NumericContentControl.docx");
    }
}
