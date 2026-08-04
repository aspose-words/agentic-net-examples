using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field that accepts a date.
        // - Name: "MyDateField"
        // - Type: Date (allows only valid date values)
        // - Format: custom date format "dd/MM/yyyy"
        // - Placeholder text: empty
        // - MaxLength: 0 (no length limit)
        FormField dateField = builder.InsertTextInput(
            "MyDateField",
            TextFormFieldType.Date,
            "dd/MM/yyyy",
            "",
            0);

        // Set the default value of the field to the current date.
        // SetTextInputValue applies the format specified above.
        dateField.SetTextInputValue(DateTime.Now);

        // Save the document to disk.
        doc.Save("FormWithDateField.docx");
    }
}
