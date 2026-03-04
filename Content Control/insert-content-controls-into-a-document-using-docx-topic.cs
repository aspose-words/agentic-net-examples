using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (SDT) at the current cursor position.
        StructuredDocumentTag plain = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        plain.Title = "PlainTextControl";
        plain.Tag = "PlainTag";
        builder.Writeln("This text is inside a plain‑text content control.");

        // Insert a rich‑text content control.
        StructuredDocumentTag rich = builder.InsertStructuredDocumentTag(SdtType.RichText);
        rich.Title = "RichTextControl";
        rich.Tag = "RichTag";
        builder.Writeln("This text is inside a rich‑text content control.");

        // Insert a checkbox content control and set its default state to checked.
        StructuredDocumentTag checkBox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkBox.Title = "CheckBoxControl";
        checkBox.Tag = "CheckTag";
        checkBox.Checked = true;
        builder.Writeln("This is a checkbox content control.");

        // Insert a drop‑down list content control and populate it with items.
        StructuredDocumentTag dropDown = builder.InsertStructuredDocumentTag(SdtType.DropDownList);
        dropDown.Title = "DropDownControl";
        dropDown.Tag = "DropDownTag";
        dropDown.ListItems.Add(new SdtListItem("Option 1", "1"));
        dropDown.ListItems.Add(new SdtListItem("Option 2", "2"));
        dropDown.ListItems.Add(new SdtListItem("Option 3", "3"));
        builder.Writeln("This is a drop‑down list content control.");

        // Insert a date picker content control and set its display format.
        StructuredDocumentTag date = builder.InsertStructuredDocumentTag(SdtType.Date);
        date.Title = "DateControl";
        date.Tag = "DateTag";
        date.DateDisplayFormat = "yyyy-MM-dd";
        builder.Writeln("This is a date picker content control.");

        // Save the document to a DOCX file.
        doc.Save("ContentControls.docx");
    }
}
