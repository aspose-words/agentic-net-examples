using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertFormFieldsExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        // The Document constructor handles opening the file.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder tied to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert a check box form field.
        // Parameters: name, defaultChecked, size (points).
        // The field is inserted at the current cursor position.
        // -------------------------------------------------
        builder.Writeln("Accept terms and conditions:");
        builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.Writeln(); // Add a paragraph break after the checkbox.

        // -------------------------------------------------
        // Insert a combo box (drop‑down) form field.
        // Parameters: name, list of items, selected index.
        // -------------------------------------------------
        builder.Writeln("Select your country:");
        string[] countries = { "USA", "Canada", "UK", "Australia", "Other" };
        builder.InsertComboBox("CountryCombo", countries, 0);
        builder.Writeln();

        // -------------------------------------------------
        // Insert a text input form field.
        // Parameters: name, type (regular text), default text, placeholder, max length.
        // -------------------------------------------------
        builder.Writeln("Enter your full name:");
        builder.InsertTextInput(
            "FullName",                     // field name
            TextFormFieldType.Regular,      // allows any text
            "",                             // default text (empty)
            "John Doe",                     // placeholder text shown in the field
            100);                           // maximum number of characters
        builder.Writeln();

        // Update all fields so that the document reflects the latest state.
        doc.UpdateFields();

        // Save the modified document to a new file.
        // The Save method determines the format from the file extension.
        doc.Save("OutputDocument.docx");
    }
}
