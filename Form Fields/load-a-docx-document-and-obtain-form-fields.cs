using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from the file system.
        // This uses the Document(string) constructor as defined in the provided rules.
        Document doc = new Document("input.docx");

        // Retrieve the collection of all form fields in the document.
        // The Range.FormFields property returns a FormFieldCollection.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Iterate through each form field and output its basic information.
        foreach (FormField field in formFields)
        {
            // Field.Name – the name of the form field.
            // Field.Type – the type of the form field (e.g., TextInput, CheckBox, DropDown).
            // Field.Result – the current displayed value of the field.
            Console.WriteLine($"Name: {field.Name}, Type: {field.Type}, Result: {field.Result}");
        }
    }
}
