using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Fields;
using Newtonsoft.Json; // Required by the task, even if not used.

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an explanatory paragraph.
        builder.Writeln("Please enter a numeric value:");

        // Insert an inline plain‑text content control (SDT) at the current cursor position.
        StructuredDocumentTag numericSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        numericSdt.Title = "NumericInput";
        numericSdt.Tag = "numeric-input";
        // Prevent the user from deleting the content control itself.
        numericSdt.LockContentControl = true;
        // Allow the user to edit the contents (the numeric validation is handled by the form field).
        numericSdt.LockContents = false;
        // Single‑line input.
        numericSdt.Multiline = false;

        // Insert a text input form field that only accepts numbers inside the SDT.
        // The overload uses 'fieldValue' for the default text.
        builder.InsertTextInput(
            name: "NumericInputField",
            type: TextFormFieldType.Number,
            format: "",
            fieldValue: "0",
            maxLength: 10);

        // Save the resulting document.
        doc.Save("NumericContentControl.docx");
    }
}
