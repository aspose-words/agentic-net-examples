using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a prompt for the user.
            builder.Writeln("Please enter the date (dd/MM/yyyy):");

            // Insert a text input form field that accepts only dates.
            // The field uses a custom date format and will be pre‑filled with the current date.
            FormField dateField = builder.InsertTextInput(
                "DateField",                     // field name
                TextFormFieldType.Date,          // restrict input to dates
                "dd/MM/yyyy",                    // custom display format
                "",                              // initial empty value
                0);                              // no length limit

            // Set the default value to the current date.
            dateField.SetTextInputValue(DateTime.Now);

            // Verify that the field was added to the document.
            FormField retrieved = doc.Range.FormFields["DateField"];
            if (retrieved == null)
                throw new InvalidOperationException("Failed to create the date form field.");

            // Save the document.
            doc.Save("FormWithDateField.docx");
        }
    }
}
