using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field with a name and default unchecked state.
        string fieldName = "MyCheckBox";
        builder.InsertCheckBox(fieldName, false, 0);

        // Save the document that contains the form field.
        string filePath = "CheckboxForm.docx";
        doc.Save(filePath);

        // Load the document (demonstrates reading from a saved file).
        Document loadedDoc = new Document(filePath);

        // Retrieve the checkbox form field by name.
        FormField checkBoxField = loadedDoc.Range.FormFields[fieldName];
        if (checkBoxField == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // Determine whether the checkbox is checked using the Checked property.
        bool isChecked = checkBoxField.Checked;

        // Output the result.
        Console.WriteLine($"Checkbox '{fieldName}' is {(isChecked ? "checked" : "unchecked")}.");
    }
}
