using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedForm.docx");
        string reopenedPath = Path.Combine(Directory.GetCurrentDirectory(), "ReopenedForm.docx");

        // ---------- Create a document with form fields ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text input field.
        FormField textField = builder.InsertTextInput(
            "TextField",                     // name
            TextFormFieldType.Regular,       // type
            "",                              // default text (unused for Regular)
            "Enter text here",               // placeholder
            50);                             // max length

        // Check box field.
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox",                      // name
            false,                           // checked value
            50);                             // size in points

        // Combo box (drop‑down) field.
        FormField comboBox = builder.InsertComboBox(
            "ComboBox",                      // name
            new[] { "Option1", "Option2", "Option3" }, // items
            0);                              // selected index

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save(originalPath);

        // ---------- Load the saved document ----------
        Document loadedDoc = new Document(originalPath);

        // Access the form fields collection.
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // Verify that at least three form fields exist.
        if (fields == null || fields.Count < 3)
            throw new InvalidOperationException("Expected at least three form fields in the document.");

        // ---------- Verify that fields are still editable ----------
        // Text input field.
        FormField loadedText = fields["TextField"];
        if (loadedText == null)
            throw new InvalidOperationException("Text field not found.");
        if (!loadedText.Enabled)
            throw new InvalidOperationException("Text field is not enabled for editing.");
        loadedText.Result = "Updated text";

        // Check box field.
        FormField loadedCheck = fields["CheckBox"];
        if (loadedCheck == null)
            throw new InvalidOperationException("Check box not found.");
        if (!loadedCheck.Enabled)
            throw new InvalidOperationException("Check box is not enabled for editing.");
        loadedCheck.Checked = true; // set the box to checked

        // Combo box field.
        FormField loadedCombo = fields["ComboBox"];
        if (loadedCombo == null)
            throw new InvalidOperationException("Combo box not found.");
        if (!loadedCombo.Enabled)
            throw new InvalidOperationException("Combo box is not enabled for editing.");
        // Change the selected index to the second item.
        loadedCombo.DropDownSelectedIndex = 1;

        // Save the document after modifications.
        loadedDoc.Save(reopenedPath);

        // Output confirmation.
        Console.WriteLine("Form fields remain editable after protection, save, and reload.");
    }
}
