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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an explanatory paragraph.
        builder.Writeln("Enter a numeric value (only digits are allowed):");

        // Insert an inline plain‑text content control (SDT).
        StructuredDocumentTag numericSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NumericInput",
            Tag = "numeric-input",
            // Prevent the user from deleting the content control.
            LockContentControl = true
        };
        builder.InsertNode(numericSdt);

        // Position the builder inside the newly created SDT.
        builder.MoveTo(numericSdt);

        // Insert a text input form field that only accepts numbers.
        // Parameters: name, type, default text, placeholder text, max length.
        builder.InsertTextInput("NumericField", TextFormFieldType.Number, "", "0", 10);

        // Save the resulting document.
        doc.Save("NumericContentControl.docx");
    }
}
