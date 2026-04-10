using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Document with Content Controls");

        // Insert a plain‑text content control (inline).
        StructuredDocumentTag plainTextSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        plainTextSdt.Title = "PlainTextControl";
        plainTextSdt.Tag = "PlainTextTag";
        builder.Write("Enter plain text here");

        // Move to a new paragraph.
        builder.Writeln();

        // Insert a rich‑text content control (block level).
        // Move the cursor to a new paragraph to ensure block level insertion.
        builder.Writeln("Rich Text Control:");
        StructuredDocumentTag richTextSdt = builder.InsertStructuredDocumentTag(SdtType.RichText);
        richTextSdt.Title = "RichTextControl";
        richTextSdt.Tag = "RichTextTag";
        // Add some formatted text inside the rich‑text control.
        builder.Font.Bold = true;
        builder.Write("Bold text inside rich‑text control. ");
        builder.Font.Italic = true;
        builder.Write("Italic text inside rich‑text control.");

        // Move to a new paragraph.
        builder.Writeln();

        // Insert a checkbox content control (inline).
        StructuredDocumentTag checkBoxSdt = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkBoxSdt.Title = "CheckBoxControl";
        checkBoxSdt.Tag = "CheckBoxTag";
        checkBoxSdt.Checked = true; // Set the default state.
        builder.Write("Checked by default");

        // Save the document to XPS format, preserving the content control boundaries.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        // No special options are required; the default behavior keeps the SDT boundaries.
        doc.Save("ContentControlsOutput.xps", xpsOptions);
    }
}
