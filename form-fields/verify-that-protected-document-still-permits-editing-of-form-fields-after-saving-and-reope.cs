using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        string filePath = Path.Combine(Environment.CurrentDirectory, "ProtectedFormFields.docx");

        // -------------------------------------------------
        // 1. Create a new document and add form fields.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text input form field.
        FormField textField = builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 0);
        textField.Result = "Initial value";
        builder.InsertParagraph();

        // Check box form field.
        FormField checkBox = builder.InsertCheckBox("CheckBox", false, 0);
        builder.InsertParagraph();

        // Combo box (drop‑down) form field.
        string[] items = { "Option1", "Option2", "Option3" };
        FormField comboBox = builder.InsertComboBox("ComboBox", items, 0);
        builder.InsertParagraph();

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save(filePath);

        // -------------------------------------------------
        // 2. Load the saved document and verify editability.
        // -------------------------------------------------
        Document loadedDoc = new Document(filePath);
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // Ensure that all expected form fields exist.
        if (fields.Count < 3)
            throw new InvalidOperationException("The document does not contain the expected form fields.");

        // Retrieve fields by name (null‑check for safety).
        FormField loadedText = fields["TextField"];
        FormField loadedCheck = fields["CheckBox"];
        FormField loadedCombo = fields["ComboBox"];

        if (loadedText == null || loadedCheck == null || loadedCombo == null)
            throw new InvalidOperationException("One or more form fields could not be found by name.");

        // Verify that the fields are enabled (editable).
        if (!loadedText.Enabled || !loadedCheck.Enabled || !loadedCombo.Enabled)
            throw new InvalidOperationException("A form field is not enabled for editing.");

        // -------------------------------------------------
        // 3. Edit the form fields after reopening.
        // -------------------------------------------------
        loadedText.Result = "Edited value";
        loadedCheck.Checked = true;
        loadedCombo.DropDownSelectedIndex = 2; // Select "Option3"

        // Validate that the changes were applied.
        if (loadedText.Result != "Edited value")
            throw new InvalidOperationException("Text field value was not updated correctly.");

        if (!loadedCheck.Checked)
            throw new InvalidOperationException("Check box value was not updated correctly.");

        if (loadedCombo.DropDownSelectedIndex != 2)
            throw new InvalidOperationException("Combo box selection was not updated correctly.");

        // Optional: save the edited document to demonstrate persistence.
        string editedPath = Path.Combine(Environment.CurrentDirectory, "EditedProtectedFormFields.docx");
        loadedDoc.Save(editedPath);

        Console.WriteLine("Form fields remain editable after protection, saving, and reopening.");
    }
}
