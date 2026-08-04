using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (unused for Regular)
            "John Doe",                      // placeholder text
            50);                             // max length

        // Insert a checkbox form field.
        builder.Writeln();
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "AcceptCheck",   // field name
            false,           // default checked state
            50);             // size in points

        // Insert a combo box (drop‑down) form field.
        builder.Writeln();
        builder.Write("Select country: ");
        string[] countries = { "USA", "Canada", "UK" };
        FormField comboBox = builder.InsertComboBox(
            "CountryBox",    // field name
            countries,       // items
            0);              // initially selected index

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        string filePath = Path.Combine(Environment.CurrentDirectory, "ProtectedFormFields.docx");
        doc.Save(filePath);

        // Load the saved document.
        Document loadedDoc = new Document(filePath);

        // ----- Verify and edit the text input field -----
        FormField loadedText = loadedDoc.Range.FormFields["NameField"];
        if (loadedText == null)
            throw new Exception("Text input field 'NameField' not found.");
        if (!loadedText.Enabled)
            throw new Exception("Text input field is not enabled.");

        // Change the field's result and validate.
        loadedText.Result = "Jane Smith";
        if (loadedText.Result != "Jane Smith")
            throw new Exception("Failed to update text input field.");

        // ----- Verify and edit the checkbox field -----
        FormField loadedCheck = loadedDoc.Range.FormFields["AcceptCheck"];
        if (loadedCheck == null)
            throw new Exception("Checkbox field 'AcceptCheck' not found.");
        if (!loadedCheck.Enabled)
            throw new Exception("Checkbox field is not enabled.");

        // Set the checkbox as checked and validate.
        loadedCheck.Checked = true;
        if (!loadedCheck.Checked)
            throw new Exception("Failed to check the checkbox field.");

        // ----- Verify and edit the combo box field -----
        FormField loadedCombo = loadedDoc.Range.FormFields["CountryBox"];
        if (loadedCombo == null)
            throw new Exception("Combo box field 'CountryBox' not found.");
        if (!loadedCombo.Enabled)
            throw new Exception("Combo box field is not enabled.");

        // Change the selected item by setting the result string.
        loadedCombo.Result = "Canada";
        if (loadedCombo.Result != "Canada")
            throw new Exception("Failed to update combo box field.");

        // Save the edited document to confirm changes persist.
        string editedPath = Path.Combine(Environment.CurrentDirectory, "EditedProtectedFormFields.docx");
        loadedDoc.Save(editedPath);

        // Indicate successful verification.
        Console.WriteLine("Form fields remain editable after protection, save, and reload.");
    }
}
