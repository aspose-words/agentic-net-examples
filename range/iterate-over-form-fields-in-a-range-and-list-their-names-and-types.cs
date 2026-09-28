using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        // Parameters: name, type, format, default text, max length.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox1", true, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field.
        // Parameters: name, list of items, selected index.
        string[] items = new string[] { "Choice1", "Choice2" };
        FormField dropdown = builder.InsertComboBox("DropDown1", items, 0);
        builder.Writeln();

        // Save the document (optional, demonstrates that the document is valid).
        doc.Save("SampleFormFields.docx");

        // Iterate over all form fields in the document's range and list their names and types.
        foreach (FormField field in doc.Range.FormFields)
        {
            Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }
    }
}
