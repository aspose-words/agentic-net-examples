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

        // Write a label before the form field.
        builder.Write("Enter date (dd/MM/yyyy): ");

        // Insert a text input form field that accepts a date.
        // - Name: "DateField"
        // - Type: Date (allows only valid dates)
        // - Format: custom date format "dd/MM/yyyy"
        // - Placeholder text: empty
        // - MaxLength: 0 (no length limit)
        FormField dateField = builder.InsertTextInput(
            "DateField",
            TextFormFieldType.Date,
            "dd/MM/yyyy",
            "",
            0);

        // Set the default value of the field to the current date.
        dateField.SetTextInputValue(DateTime.Now);

        // Save the document to a file in the current directory.
        doc.Save("FormWithDate.docx");
    }
}
