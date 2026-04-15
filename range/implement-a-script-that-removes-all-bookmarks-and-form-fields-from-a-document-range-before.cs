using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content with bookmarks.
        builder.StartBookmark("SampleBookmark");
        builder.Write("This text is inside a bookmark.");
        builder.EndBookmark("SampleBookmark");
        builder.Writeln();

        // Insert a few form fields.
        builder.InsertComboBox("ComboBoxField", new[] { "Option1", "Option2", "Option3" }, 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBoxField", false, 50);
        builder.Writeln();
        builder.InsertTextInput("TextInputField", TextFormFieldType.Regular, "", "Placeholder text", 50);
        builder.Writeln();

        // At this point the document contains bookmarks and form fields.
        // Remove all bookmarks from the whole document range.
        doc.Range.Bookmarks.Clear();

        // Remove all form fields from the whole document range.
        doc.Range.FormFields.Clear();

        // Save the cleaned document.
        doc.Save("CleanedDocument.docx");
    }
}
