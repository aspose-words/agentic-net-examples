using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "John Doe", 50);

        // Insert a checkbox form field.
        builder.Writeln();
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("CheckBoxField", false, 50);

        // Insert a combo box (dropdown) form field.
        builder.Writeln();
        builder.Write("Select a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("ComboBoxField", items, 0);

        // Save the document as required.
        doc.Save("FormFields.docx");

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate over each form field and output its name and type.
        foreach (FormField field in formFields)
        {
            if (field != null)
                Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }
    }
}
