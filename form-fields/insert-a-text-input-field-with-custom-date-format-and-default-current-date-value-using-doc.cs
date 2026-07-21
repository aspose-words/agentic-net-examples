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

        // Add a label before the form field.
        builder.Writeln("Enter date (dd/MM/yyyy):");

        // Insert a text input form field that only accepts dates.
        // - Name: "DateField"
        // - Type: Date (restricts input to valid dates)
        // - Format: custom date format "dd/MM/yyyy"
        // - Initial displayed text: empty
        // - MaxLength: 0 (no length limit)
        FormField dateField = builder.InsertTextInput(
            "DateField",
            TextFormFieldType.Date,
            "dd/MM/yyyy",
            "",
            0);

        // Set the default value of the field to the current date.
        dateField.SetTextInputValue(DateTime.Now);

        // Save the document to the local file system.
        doc.Save("FormFieldDate.docx");
    }
}
