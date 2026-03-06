using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (SDT) and set its metadata.
        StructuredDocumentTag plain = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        plain.Title = "PlainTextControl";
        plain.Tag = "PlainTag";
        builder.Writeln("Plain text inside control.");

        // Insert a rich‑text content control.
        StructuredDocumentTag rich = builder.InsertStructuredDocumentTag(SdtType.RichText);
        rich.Title = "RichTextControl";
        rich.Tag = "RichTag";
        builder.Writeln("Rich text inside control.");

        // Insert a checkbox content control and set it to checked.
        StructuredDocumentTag checkbox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkbox.Title = "CheckBoxControl";
        checkbox.Tag = "CheckTag";
        checkbox.Checked = true;
        builder.Writeln("Checkbox control.");

        // Insert a drop‑down list content control and populate its items.
        StructuredDocumentTag dropdown = builder.InsertStructuredDocumentTag(SdtType.DropDownList);
        dropdown.Title = "DropDownControl";
        dropdown.Tag = "DropTag";
        dropdown.ListItems.Add(new SdtListItem("Option 1", "1"));
        dropdown.ListItems.Add(new SdtListItem("Option 2", "2"));
        dropdown.ListItems.Add(new SdtListItem("Option 3", "3"));
        builder.Writeln("Dropdown control.");

        // Save the document as a DOCX file.
        doc.Save("ContentControls.docx");
    }
}
