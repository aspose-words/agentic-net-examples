using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsRangeCleanup
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bookmark with some text.
            builder.StartBookmark("SampleBookmark");
            builder.Write("This text is inside a bookmark.");
            builder.EndBookmark("SampleBookmark");
            builder.Writeln();

            // Insert a check box form field.
            builder.InsertCheckBox("CheckBoxField", false, 50);
            builder.Writeln();

            // Insert a combo box form field.
            builder.InsertComboBox("ComboBoxField", new[] { "Option1", "Option2", "Option3" }, 0);
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
}
