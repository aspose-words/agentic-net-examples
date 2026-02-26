using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertFormFieldsExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        // The Document constructor handles opening the file.
        Document doc = new Document("InputDocument.docx");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document so that new content is appended.
        builder.MoveToDocumentEnd();

        // Insert a paragraph break before adding form fields for readability.
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a CheckBox form field.
        // Parameters: name, defaultChecked, size (points).
        // The field will be placed at the current cursor position.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 50);
        // Optional: configure additional properties.
        checkBox.HelpText = "Check to accept the terms.";
        checkBox.OwnHelp = true;

        // Insert a line break between fields.
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a ComboBox (drop‑down) form field.
        // Parameters: name, list of items, selected index.
        string[] colors = { "Red", "Green", "Blue", "Yellow" };
        builder.Write("Select a color: ");
        FormField comboBox = builder.InsertComboBox("ColorChoice", colors, 0);
        comboBox.CalculateOnExit = true; // Recalculate dependent fields when selection changes.

        // Insert a line break between fields.
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a TextInput form field.
        // Parameters: name, type (regular text), default text, placeholder, max length.
        builder.Write("Enter your name: ");
        FormField textInput = builder.InsertTextInput(
            "UserName",
            TextFormFieldType.Regular,
            "",               // No default text.
            "John Doe",       // Placeholder shown until the user types.
            50);              // Maximum length of the input.

        // Optional: set macros or formatting if needed.
        textInput.EntryMacro = "OnEnterName";
        textInput.ExitMacro = "OnExitName";

        // -------------------------------------------------
        // Update all fields so that any calculated results are reflected.
        doc.UpdateFields();

        // Save the modified document to a new file.
        doc.Save("OutputDocumentWithFormFields.docx");
    }
}
