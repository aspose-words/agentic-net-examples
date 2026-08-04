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

        // Add a paragraph of normal text – this will be read‑only after protection.
        builder.Writeln("This part of the document is read‑only.");

        // Insert a text input form field.
        FormField textField = builder.InsertTextInput(
            "TextInput1",                     // field name
            TextFormFieldType.Regular,        // field type
            "",                               // default text (empty)
            "Enter name",                     // placeholder text
            0);                               // max length (0 = unlimited)
        // Set an initial value for the text field.
        textField.Result = "John Doe";

        // Insert a check box form field.
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox1",   // field name
            false,         // default unchecked
            15);           // size in points

        // Insert a combo box (drop‑down) form field.
        string[] items = { "Option A", "Option B", "Option C" };
        FormField comboBox = builder.InsertComboBox(
            "ComboBox1",   // field name
            items,         // list items
            0);            // initially selected index

        // Validate that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
            throw new InvalidOperationException("No form fields were created.");

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document to the current directory.
        doc.Save("ProtectedFormFields.docx");
    }
}
