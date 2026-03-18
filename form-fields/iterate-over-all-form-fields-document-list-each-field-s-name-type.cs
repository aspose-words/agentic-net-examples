using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class ListFormFields
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Get the collection of all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Iterate over the collection using its enumerator.
        using (IEnumerator<FormField> enumerator = formFields.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                FormField field = enumerator.Current;
                // Output the field's name and its type.
                Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
            }
        }

        // Save the document (unchanged) to a new file.
        doc.Save("Output.docx");
    }
}
