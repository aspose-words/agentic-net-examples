using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        // Parameters: name, defaultValue, checkedValue, size (points).
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, true, 30);

        // Enable explicit size handling for the checkbox.
        checkBox.IsCheckBoxExactSize = true;

        // Set a custom size (in points). This takes effect because IsCheckBoxExactSize is true.
        checkBox.CheckBoxSize = 30;

        // Ensure the checkbox is checked by default.
        checkBox.Checked = true;
        checkBox.Default = true;

        // Save the document to a file.
        doc.Save("CheckboxFormField.docx");
    }
}
