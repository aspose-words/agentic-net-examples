using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field that automatically contains the current date.
        // TextFormFieldType.CurrentDate makes the field value the current date.
        // The third argument is the format string that defines how the date will be displayed.
        // The field value argument is left empty because the field generates its own value.
        // maxLength = 0 means there is no length restriction (irrelevant for a date field).
        FormField dateField = builder.InsertTextInput(
            "MyDateField",                // field name (also creates a bookmark)
            TextFormFieldType.CurrentDate, // field type – current date
            "dd MMMM yyyy",               // custom display format
            "",                           // initial value (ignored for CurrentDate)
            0);                           // unlimited length

        // Update all fields so the current date appears immediately.
        doc.UpdateFields();

        // Save the document.
        doc.Save("DateInputForm.docx");
    }
}
