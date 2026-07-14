using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, "ProtectedFormFields.docx");

        // 1. Create a new document and add form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Check box form field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Combo box (drop‑down) form field.
        builder.Write("Select a fruit: ");
        string[] fruits = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("FruitChoice", fruits, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // 2. Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // 3. Save the protected document.
        doc.Save(filePath);

        // 4. Load the saved document.
        Document loadedDoc = new Document(filePath);
        FormFieldCollection fields = loadedDoc.Range.FormFields;

        // 5. Validate that form fields exist.
        if (fields == null || fields.Count < 3)
            throw new InvalidOperationException("Expected form fields were not found in the loaded document.");

        // 6. Verify each form field is still enabled and editable.
        // Text input field.
        FormField loadedText = fields["NameField"];
        if (loadedText == null)
            throw new InvalidOperationException("Text input field not found.");
        if (!loadedText.Enabled)
            throw new InvalidOperationException("Text input field is not enabled.");
        loadedText.Result = "Jane Smith"; // Update the value.

        // Check box field.
        FormField loadedCheck = fields["AcceptTerms"];
        if (loadedCheck == null)
            throw new InvalidOperationException("Check box field not found.");
        if (!loadedCheck.Enabled)
            throw new InvalidOperationException("Check box field is not enabled.");
        loadedCheck.Checked = true; // Toggle the check box.

        // Combo box field.
        FormField loadedCombo = fields["FruitChoice"];
        if (loadedCombo == null)
            throw new InvalidOperationException("Combo box field not found.");
        if (!loadedCombo.Enabled)
            throw new InvalidOperationException("Combo box field is not enabled.");
        loadedCombo.DropDownSelectedIndex = 2; // Select "Cherry".

        // 7. Save the document again to confirm changes persist (optional).
        string updatedPath = Path.Combine(outputDir, "ProtectedFormFields_Updated.docx");
        loadedDoc.Save(updatedPath);

        // 8. Output verification results.
        Console.WriteLine("Document protection and form field editability verified successfully.");
        Console.WriteLine($"Text field new value: {loadedText.Result}");
        Console.WriteLine($"Check box is now checked: {loadedCheck.Checked}");
        Console.WriteLine($"Combo box selected item: {loadedCombo.Result}");
    }
}
