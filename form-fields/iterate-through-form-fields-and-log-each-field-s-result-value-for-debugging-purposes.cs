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

        // Insert a text input form field.
        FormField textField = builder.InsertTextInput(
            "TextField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (empty)
            "Placeholder text",              // placeholder text
            50);                             // maximum length

        // Insert a checkbox form field.
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox",                      // field name
            false,                           // default unchecked
            50);                             // size in points

        // Insert a combo box (dropdown) form field.
        FormField comboBox = builder.InsertComboBox(
            "DropDown",                      // field name
            new[] { "Option1", "Option2", "Option3" }, // items
            0);                              // default selected index

        // Set some example values.
        textField.Result = "Sample text";
        checkBox.Checked = true;               // will affect Result as "1"
        comboBox.DropDownSelectedIndex = 2;    // selects "Option3"

        // Access the collection of all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate through each form field and log its Result value.
        foreach (FormField field in formFields)
        {
            // For debugging purposes we output the field name and its current result.
            Console.WriteLine($"Field '{field.Name}' Result: {field.Result}");
        }

        // Save the document to disk.
        doc.Save("FormFieldsOutput.docx");
    }
}
