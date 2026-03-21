using System;
using Aspose.Words;
using Aspose.Words.Fields;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Create a new document and add a couple of form fields for demonstration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Sample text", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", false, 0);
        builder.Writeln();

        // Insert a dropdown form field.
        builder.InsertComboBox("DropDownField", new[] { "Option1", "Option2", "Option3" }, 0);
        builder.Writeln();

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

        // Save the document (optional – the document itself is unchanged except for the added fields).
        doc.Save("FormFields_Logged.docx");
    }
}
