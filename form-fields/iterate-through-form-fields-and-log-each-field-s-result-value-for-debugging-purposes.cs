using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder to add form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field. The last parameter is the maximum length of the field.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default Text", 0);

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);

        // Insert a dropdown (combo box) form field.
        builder.InsertComboBox("DropDown1", new string[] { "Option1", "Option2", "Option3" }, 0);

        // Save the document that now contains the form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Ensure that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Iterate through all form fields and log their values.
        foreach (FormField field in doc.Range.FormFields)
        {
            if (field == null)
                continue; // Guard against null, though collection should not contain nulls.

            string fieldInfo;

            switch (field.Type)
            {
                case FieldType.FieldFormCheckBox:
                    fieldInfo = $"Name: {field.Name}, Type: CheckBox, Checked: {field.Checked}";
                    break;
                case FieldType.FieldFormDropDown:
                    fieldInfo = $"Name: {field.Name}, Type: DropDown, Selected: {field.Result}";
                    break;
                case FieldType.FieldFormTextInput:
                    fieldInfo = $"Name: {field.Name}, Type: TextInput, Result: {field.Result}";
                    break;
                default:
                    fieldInfo = $"Name: {field.Name}, Type: Unknown, Result: {field.Result}";
                    break;
            }

            Console.WriteLine(fieldInfo);
        }

        // Save the document again (even though we didn't modify fields, this satisfies the save requirement).
        const string finalPath = "FormFields_Output.docx";
        doc.Save(finalPath);
    }
}
