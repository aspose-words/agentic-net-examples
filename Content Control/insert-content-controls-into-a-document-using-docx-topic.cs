using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertContentControls
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Document with various content controls (Structured Document Tags):");
        builder.Writeln();

        // ---------- Plain Text Content Control ----------
        // Insert a plain‑text content control.
        StructuredDocumentTag plainText = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        // Set a title and placeholder text.
        plainText.Title = "PlainTextControl";
        plainText.PlaceholderName = "Enter plain text here";
        // Write sample text inside the control.
        builder.Writeln("Sample plain text inside the control.");
        // Move the cursor out of the control.
        builder.MoveToDocumentEnd();

        // ---------- Rich Text Content Control ----------
        builder.Writeln();
        StructuredDocumentTag richText = builder.InsertStructuredDocumentTag(SdtType.RichText);
        richText.Title = "RichTextControl";
        richText.PlaceholderName = "Enter rich text here";
        // Apply some formatting inside the rich‑text control.
        builder.Font.Bold = true;
        builder.Font.Size = 14;
        builder.Writeln("Bold rich text inside the control.");
        // Reset formatting.
        builder.Font.ClearFormatting();
        builder.MoveToDocumentEnd();

        // ---------- Checkbox Content Control ----------
        builder.Writeln();
        StructuredDocumentTag checkBox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkBox.Title = "AgreementCheckBox";
        // Set the initial state of the checkbox.
        checkBox.Checked = true;
        builder.Writeln("I agree to the terms and conditions.");
        builder.MoveToDocumentEnd();

        // ---------- Drop‑Down List Content Control ----------
        builder.Writeln();
        StructuredDocumentTag dropDown = builder.InsertStructuredDocumentTag(SdtType.DropDownList);
        dropDown.Title = "CountrySelector";
        // Add items to the drop‑down list.
        dropDown.ListItems.Add(new SdtListItem("United States"));
        dropDown.ListItems.Add(new SdtListItem("Canada"));
        dropDown.ListItems.Add(new SdtListItem("United Kingdom"));
        // Write a label for the control.
        builder.Writeln("Select a country:");
        builder.MoveToDocumentEnd();

        // ---------- Date Picker Content Control ----------
        builder.Writeln();
        StructuredDocumentTag datePicker = builder.InsertStructuredDocumentTag(SdtType.Date);
        datePicker.Title = "BirthDatePicker";
        // Set display format (e.g., "MM/dd/yyyy").
        datePicker.DateDisplayFormat = "MM/dd/yyyy";
        // Write a label.
        builder.Writeln("Select your birth date:");
        builder.MoveToDocumentEnd();

        // ---------- Picture Content Control ----------
        builder.Writeln();
        StructuredDocumentTag pictureControl = builder.InsertStructuredDocumentTag(SdtType.Picture);
        pictureControl.Title = "ProfilePicture";
        // Write a label.
        builder.Writeln("Insert a profile picture:");
        builder.MoveToDocumentEnd();

        // Save the document to a DOCX file.
        doc.Save("ContentControls.docx");
    }
}
