using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeFormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a label before the form field.
            builder.Write("Enter date (dd/MM/yyyy): ");

            // Insert a text input form field that accepts only dates.
            // The field is given a name so we can locate it later.
            builder.InsertTextInput("DateField", TextFormFieldType.Date, "", "", 0);

            // Retrieve the inserted form field from the document's collection.
            FormField dateField = null;
            foreach (FormField ff in doc.Range.FormFields)
            {
                if (ff.Name == "DateField")
                {
                    dateField = ff;
                    break;
                }
            }

            // Validate that the field was found.
            if (dateField == null)
                throw new InvalidOperationException("The form field 'DateField' was not found.");

            // Set a custom display format for the date.
            dateField.TextInputFormat = "dd/MM/yyyy";

            // Set the default value to the current date.
            dateField.SetTextInputValue(DateTime.Now);

            // Save the document.
            doc.Save("FormFieldDate.docx");
        }
    }
}
