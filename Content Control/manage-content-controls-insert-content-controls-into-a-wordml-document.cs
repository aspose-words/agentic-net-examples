using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Create a DocumentBuilder which will be used to insert nodes.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Insert a plain‑text content control (inline).
            // -------------------------------------------------
            StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(
                doc,
                SdtType.PlainText,          // The type of the content control.
                MarkupLevel.Inline);        // Occurs at the inline level.

            plainTextSdt.Title = "Plain Text Control";
            plainTextSdt.Tag = "PlainTextTag";

            // Insert the content control into the document.
            builder.InsertNode(plainTextSdt);
            // Add some placeholder text inside the control.
            builder.Writeln("Enter plain text here...");

            // -------------------------------------------------
            // Insert a rich‑text content control (block level).
            // -------------------------------------------------
            StructuredDocumentTag richTextSdt = new StructuredDocumentTag(
                doc,
                SdtType.RichText,
                MarkupLevel.Block);

            richTextSdt.Title = "Rich Text Control";
            richTextSdt.Tag = "RichTextTag";

            // Insert a new paragraph before the block‑level control.
            builder.Writeln(); // Ensure we are on a new paragraph.
            builder.InsertNode(richTextSdt);
            // Add some formatted text inside the rich‑text control.
            builder.Font.Bold = true;
            builder.Writeln("Bold text inside rich‑text control.");
            builder.Font.Bold = false;

            // -------------------------------------------------
            // Insert a checkbox content control.
            // -------------------------------------------------
            StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(
                doc,
                SdtType.Checkbox,
                MarkupLevel.Inline);

            checkBoxSdt.Title = "Checkbox Control";
            checkBoxSdt.Tag = "CheckBoxTag";
            checkBoxSdt.Checked = false; // Default unchecked.

            builder.Writeln(); // New line before the checkbox.
            builder.InsertNode(checkBoxSdt);
            builder.Writeln("Check this box if you agree.");

            // -------------------------------------------------
            // Insert a date picker content control.
            // -------------------------------------------------
            StructuredDocumentTag dateSdt = new StructuredDocumentTag(
                doc,
                SdtType.Date,
                MarkupLevel.Inline);

            dateSdt.Title = "Date Picker Control";
            dateSdt.Tag = "DateTag";
            dateSdt.DateDisplayFormat = "MMMM d, yyyy"; // Example display format.

            builder.Writeln(); // New line before the date picker.
            builder.InsertNode(dateSdt);
            builder.Writeln("Select a date.");

            // -------------------------------------------------
            // Save the document to disk.
            // -------------------------------------------------
            doc.Save("ContentControls.docx");
        }
    }
}
