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

        // Add a regular text input form field.
        // The last parameter is maxLength (0 = no limit).
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default text", 0);

        // Add a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);

        // Add a combo box form field with two items.
        // Provide the list of items and the initially selected index.
        builder.InsertComboBox("ComboBox1", new string[] { "Option1", "Option2" }, 0);

        // Save the document (optional, demonstrates that the document can be persisted).
        doc.Save("FormFields.docx");

        // Iterate over all form fields in the document's range and output their name and type.
        foreach (FormField field in doc.Range.FormFields)
        {
            Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }
    }
}
