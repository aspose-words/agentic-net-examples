using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FormFieldInsertion
{
    static void Main()
    {
        // Load an existing DOCX file.
        // The Document constructor is the approved way to create/load a document.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document where we will add the form fields.
        builder.MoveToDocumentEnd();

        // Insert a paragraph break before the first form field.
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a CheckBox form field.
        // Parameters: name, default checked state, size (points).
        // Example: a checkbox named "AcceptTerms" that is unchecked by default and 50 points wide.
        builder.Write("Accept terms and conditions: ");
        builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a ComboBox (drop‑down) form field.
        // Parameters: name, array of items, selected index.
        // Example: a combo box named "CountrySelect" with three options.
        string[] countries = { "USA", "Canada", "Mexico" };
        builder.Write("Select your country: ");
        builder.InsertComboBox("CountrySelect", countries, 0);
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a TextInput form field.
        // Parameters: name, type (regular text), default text, placeholder, max length.
        // Example: a text input named "UserName" with a placeholder.
        builder.Write("Enter your name: ");
        builder.InsertTextInput(
            "UserName",                     // field name
            TextFormFieldType.Regular,     // field type
            "",                            // default text (empty)
            "John Doe",                    // placeholder text
            50);                           // maximum length

        // -------------------------------------------------
        // After inserting form fields, update all fields so that any dependent fields reflect the changes.
        doc.UpdateFields();

        // Save the modified document.
        // The Save method is the approved way to persist the document.
        doc.Save("OutputDocument.docx");
    }
}
