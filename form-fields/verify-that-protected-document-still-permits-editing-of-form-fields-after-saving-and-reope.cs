using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        const string filePath = "FormFields.docx";
        const string updatedFilePath = "FormFields_Updated.docx";

        // Create a new document and add form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.InsertParagraph();
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default Text", 0);
        builder.Writeln();

        builder.InsertCheckBox("CheckBoxField", false, 0);
        builder.Writeln();

        builder.InsertComboBox("ComboBoxField", new[] { "Option1", "Option2", "Option3" }, 0);
        builder.Writeln();

        doc.Save(filePath);

        // Reopen the saved document.
        Document loadedDoc = new Document(filePath);
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // Validate that each expected form field exists.
        FormField textField = fields["TextField"];
        FormField checkBoxField = fields["CheckBoxField"];
        FormField comboBoxField = fields["ComboBoxField"];

        if (textField == null) throw new InvalidOperationException("TextField not found.");
        if (checkBoxField == null) throw new InvalidOperationException("CheckBoxField not found.");
        if (comboBoxField == null) throw new InvalidOperationException("ComboBoxField not found.");

        // Update the form fields.
        textField.Result = "Updated Text";
        checkBoxField.Checked = true;
        comboBoxField.Result = "Option2";

        // Verify that updates were applied.
        if (textField.Result != "Updated Text") throw new InvalidOperationException("TextField update failed.");
        if (!checkBoxField.Checked) throw new InvalidOperationException("CheckBoxField update failed.");
        if (comboBoxField.Result != "Option2") throw new InvalidOperationException("ComboBoxField update failed.");

        // Save the modified document.
        loadedDoc.Save(updatedFilePath);
    }
}
