using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertContentControls
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert a plain‑text content control (inline).
        // -------------------------------------------------
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(
            doc,                     // The document to which the SDT belongs.
            SdtType.PlainText,      // Type of the content control.
            MarkupLevel.Inline);    // Inline placement.

        // Set optional properties.
        plainTextSdt.Title = "Plain Text Control";
        plainTextSdt.Tag = "PlainTextTag";

        // Insert the SDT into the document.
        builder.InsertNode(plainTextSdt);

        // Move the cursor inside the SDT to add placeholder text.
        builder.MoveTo(plainTextSdt);
        builder.Write("Enter plain text here");

        // Add a paragraph break after the control.
        builder.Writeln();

        // -------------------------------------------------
        // 2. Insert a rich‑text content control (block level).
        // -------------------------------------------------
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(
            doc,
            SdtType.RichText,
            MarkupLevel.Block);

        richTextSdt.Title = "Rich Text Control";
        richTextSdt.Tag = "RichTextTag";

        // Insert the SDT.
        builder.InsertNode(richTextSdt);
        builder.Writeln(); // Ensure the SDT starts on a new line.

        // Move inside the SDT and add formatted content.
        builder.MoveTo(richTextSdt);
        builder.Font.Bold = true;
        builder.Font.Size = 14;
        builder.Write("Bold rich text");
        builder.Font.Bold = false;
        builder.Writeln();
        builder.Font.Italic = true;
        builder.Write("Italic rich text");
        builder.Font.Italic = false;

        // End the block‑level SDT with a paragraph break.
        builder.Writeln();

        // -------------------------------------------------
        // 3. Insert a checkbox content control.
        // -------------------------------------------------
        // NOTE: The CheckBox SDT type is available only in Aspose.Words versions 22.5 and later.
        // If you are using an older version, either upgrade the library or comment out this block.
#if ASPOSE_WORDS_SUPPORTS_CHECKBOX
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(
            doc,
            SdtType.CheckBox,
            MarkupLevel.Inline);

        checkBoxSdt.Title = "Agreement Checkbox";
        checkBoxSdt.Tag = "AgreeCheck";
        checkBoxSdt.Checked = false; // Default unchecked.

        // Insert the checkbox SDT.
        builder.InsertNode(checkBoxSdt);
        builder.Writeln(); // Separate from following text.
#endif

        // -------------------------------------------------
        // 4. Insert a date picker content control.
        // -------------------------------------------------
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(
            doc,
            SdtType.Date,
            MarkupLevel.Inline);

        dateSdt.Title = "Date Picker";
        dateSdt.Tag = "DateTag";
        dateSdt.DateDisplayFormat = "MMMM d, yyyy"; // e.g., "January 1, 2023"

        // Insert the date SDT.
        builder.InsertNode(dateSdt);
        builder.Writeln();

        // -------------------------------------------------
        // 5. Insert a dropdown list content control.
        // -------------------------------------------------
        StructuredDocumentTag comboBoxSdt = new StructuredDocumentTag(
            doc,
            SdtType.ComboBox,
            MarkupLevel.Inline);

        comboBoxSdt.Title = "Country Selector";
        comboBoxSdt.Tag = "CountryTag";

        // Populate the list items.
        comboBoxSdt.ListItems.Add(new SdtListItem("United States", "US"));
        comboBoxSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
        comboBoxSdt.ListItems.Add(new SdtListItem("United Kingdom", "UK"));
        comboBoxSdt.ListItems.Add(new SdtListItem("Australia", "AU"));

        // Insert the dropdown SDT.
        builder.InsertNode(comboBoxSdt);
        builder.Writeln();

        // -------------------------------------------------
        // Save the document to a DOCX file.
        // -------------------------------------------------
        doc.Save("ContentControls.docx");
    }
}
