using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary documents.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedForm.docx");
        string reopenedPath = Path.Combine(Directory.GetCurrentDirectory(), "ReopenedForm.docx");

        // 1. Create a new document and add form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text input field.
        FormField textField = builder.InsertTextInput(
            "TextField",
            TextFormFieldType.Regular,
            "",
            "Enter text",
            50);
        textField.Enabled = true;

        // Check box field.
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox",
            false,
            50);
        checkBox.Enabled = true;

        // Combo box (drop‑down) field.
        FormField comboBox = builder.InsertComboBox(
            "ComboBox",
            new[] { "Option1", "Option2", "Option3" },
            0);
        comboBox.Enabled = true;

        // 2. Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // 3. Save the protected document.
        doc.Save(originalPath);

        // 4. Load the saved document.
        Document loadedDoc = new Document(originalPath);

        // 5. Verify that form fields exist and are still editable.
        FormFieldCollection fields = loadedDoc.Range.FormFields;
        if (fields == null || fields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were found in the loaded document.");
        }

        // Iterate through each form field and perform a simple edit.
        foreach (FormField field in fields)
        {
            if (!field.Enabled)
                throw new InvalidOperationException($"Form field '{field.Name}' is not enabled.");

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Update the text field's result.
                    field.Result = "NewValue";
                    break;

                case FieldType.FieldFormCheckBox:
                    // Check the checkbox.
                    field.Checked = true;
                    break;

                case FieldType.FieldFormDropDown:
                    // Select the second item if it exists.
                    if (field.DropDownItems.Count > 1)
                        field.DropDownSelectedIndex = 1;
                    break;

                default:
                    // Other field types are not expected in this example.
                    break;
            }
        }

        // 6. Save the document after modifications.
        loadedDoc.Save(reopenedPath);

        // 7. Output verification result.
        Console.WriteLine("Protected document was saved, reopened, and form fields remain editable.");
    }
}
