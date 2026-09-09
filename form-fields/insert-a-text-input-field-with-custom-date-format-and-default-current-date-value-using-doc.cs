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

        // Add a label for the form field.
        builder.Writeln("Enter date (dd/MM/yyyy):");

        // Insert a text input form field that accepts dates.
        // - Name: "DateField"
        // - Type: Date (allows only valid date values)
        // - Format: custom date format "dd/MM/yyyy"
        // - Initial value: empty (will be set below)
        // - MaxLength: 0 (no length limit)
        FormField dateField = builder.InsertTextInput(
            "DateField",
            TextFormFieldType.Date,
            "dd/MM/yyyy",
            "",
            0);

        // Set the default value of the field to the current date.
        dateField.SetTextInputValue(DateTime.Now);

        // Save the document to a file.
        doc.Save("FormWithDateField.docx");
    }
}
