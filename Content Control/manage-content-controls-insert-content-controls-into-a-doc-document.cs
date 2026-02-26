using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document – this positions the cursor at the start.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert a plain‑text content control (inline).
        // -------------------------------------------------
        StructuredDocumentTag plainText = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        plainText.Title = "PlainTextControl";
        plainText.Tag = "PlainTextTag";
        builder.InsertNode(plainText);
        // Move the cursor inside the newly created SDT.
        builder.MoveTo(plainText.LastChild);
        builder.Writeln("This text is inside a plain‑text content control.");

        // -------------------------------------------------
        // 2. Insert a rich‑text content control (block).
        // -------------------------------------------------
        StructuredDocumentTag richText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richText.Title = "RichTextControl";
        richText.Tag = "RichTextTag";
        builder.InsertNode(richText);
        builder.MoveTo(richText.LastChild);
        builder.Writeln("This text is inside a rich‑text content control.");

        // -------------------------------------------------
        // 3. Insert a checkbox content control (inline).
        // -------------------------------------------------
        // NOTE: The CheckBox SDT type is available starting from Aspose.Words 22.5.
        // If you are using an older version, replace it with a plain‑text SDT as shown below.
        StructuredDocumentTag checkBox = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        checkBox.Title = "CheckBoxControl";
        checkBox.Tag = "CheckBoxTag";
        // Simulate a checked state by inserting a checked box character.
        builder.InsertNode(checkBox);
        builder.MoveTo(checkBox.LastChild);
        builder.Writeln("☑ Checkbox is checked.");

        // -------------------------------------------------
        // 4. Insert a combo‑box content control (inline) with items.
        // -------------------------------------------------
        StructuredDocumentTag comboBox = new StructuredDocumentTag(doc, SdtType.ComboBox, MarkupLevel.Inline);
        comboBox.Title = "ComboBoxControl";
        comboBox.Tag = "ComboBoxTag";
        // Populate the list of items.
        comboBox.ListItems.Add(new SdtListItem("Option 1", "1"));
        comboBox.ListItems.Add(new SdtListItem("Option 2", "2"));
        comboBox.ListItems.Add(new SdtListItem("Option 3", "3"));
        builder.InsertNode(comboBox);
        builder.MoveTo(comboBox.LastChild);
        builder.Writeln("Select an option from the combo box.");

        // Save the document to a file. The format is inferred from the extension.
        doc.Save("ContentControls.docx");
    }
}
