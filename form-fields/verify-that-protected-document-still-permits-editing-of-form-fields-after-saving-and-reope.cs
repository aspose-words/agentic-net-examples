using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        const string protectedPath = "ProtectedFormFields.docx";

        // Create a new document and insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text input field.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.Writeln();

        // Checkbox field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptCheck", false, 0);
        builder.Writeln();

        // Combo box field.
        builder.Write("Select country: ");
        string[] countries = { "USA", "Canada", "Mexico" };
        FormField comboBox = builder.InsertComboBox("CountryBox", countries, 0);
        builder.Writeln();

        // Protect the document to allow only form field editing.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save(protectedPath);

        // Load the saved document.
        Document loadedDoc = new Document(protectedPath);
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // Ensure at least one form field exists.
        if (fields == null || fields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        // Update the text input field.
        FormField loadedText = fields["NameField"];
        if (loadedText == null)
            throw new InvalidOperationException("Text input field 'NameField' not found.");
        loadedText.Result = "Alice";

        // Update the checkbox field.
        FormField loadedCheck = fields["AcceptCheck"];
        if (loadedCheck == null)
            throw new InvalidOperationException("Checkbox field 'AcceptCheck' not found.");
        loadedCheck.Checked = true;

        // Update the combo box field (select "Canada", index 1).
        FormField loadedCombo = fields["CountryBox"];
        if (loadedCombo == null)
            throw new InvalidOperationException("Combo box field 'CountryBox' not found.");
        loadedCombo.DropDownSelectedIndex = 1;

        // Validate that updates were applied.
        if (loadedText.Result != "Alice")
            throw new InvalidOperationException("Failed to update text input field.");
        if (!loadedCheck.Checked)
            throw new InvalidOperationException("Failed to update checkbox field.");
        if (loadedCombo.DropDownSelectedIndex != 1)
            throw new InvalidOperationException("Failed to update combo box field.");

        // Save the modified document.
        const string modifiedPath = "ModifiedFormFields.docx";
        loadedDoc.Save(modifiedPath);

        // Output the final values to the console.
        Console.WriteLine($"Text field value: {loadedText.Result}");
        Console.WriteLine($"Checkbox checked: {loadedCheck.Checked}");
        Console.WriteLine($"Combo box selected: {loadedCombo.Result}");
    }
}
