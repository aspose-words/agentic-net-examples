using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Non‑field content that should become read‑only after protection.
        string nonFieldText = "This paragraph should not be editable when the document is protected.";
        builder.Writeln(nonFieldText);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",
            TextFormFieldType.Regular,
            "",
            "John Doe",
            50);
        textField.Result = "John Doe";

        // Insert a checkbox form field.
        builder.InsertParagraph();
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptCheck", false, 0);
        checkBox.Checked = false;

        // Insert a combo box (drop‑down) form field.
        builder.InsertParagraph();
        builder.Write("Select an option: ");
        string[] items = { "Option1", "Option2", "Option3" };
        FormField comboBox = builder.InsertComboBox("OptionCombo", items, 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        const string outputPath = "ProtectedFormFields.docx";
        doc.Save(outputPath);

        // ----- Verification -----
        // 1. Verify that the document protection type is set correctly.
        if (doc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection type is not set to AllowOnlyFormFields.");

        // 2. Verify that the expected number of form fields exist.
        const int expectedFormFieldCount = 3;
        if (doc.Range.FormFields.Count != expectedFormFieldCount)
            throw new InvalidOperationException($"Expected {expectedFormFieldCount} form fields, but found {doc.Range.FormFields.Count}.");

        // 3. Verify that the non‑field paragraph is present and unchanged.
        bool paragraphFound = false;
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            string paraText = para.GetText().TrimEnd('\r', '\n');
            if (paraText == nonFieldText)
            {
                paragraphFound = true;
                break;
            }
        }

        if (!paragraphFound)
            throw new InvalidOperationException("Non‑field content was altered or missing after protection.");

        // Output verification results.
        Console.WriteLine("Document protected successfully.");
        Console.WriteLine($"Protection type: {doc.ProtectionType}");
        Console.WriteLine($"Form fields count: {doc.Range.FormFields.Count}");
        Console.WriteLine("Non‑field content remains unchanged.");
    }
}
