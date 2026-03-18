using System;
using Aspose.Words;
using Aspose.Words.Fields;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Load an existing Word document that contains form fields.
        Document doc = new Document("FormFields.docx");

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Iterate through the collection using its enumerator.
        using (IEnumerator<FormField> enumerator = formFields.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                FormField field = enumerator.Current;

                // Log the field's name and its current Result value.
                // Result holds the displayed content of the form field.
                Console.WriteLine($"Field \"{field.Name}\" Result: {field.Result}");
            }
        }

        // Save the document (optional – the document itself is unchanged).
        doc.Save("FormFields_Logged.docx");
    }
}
