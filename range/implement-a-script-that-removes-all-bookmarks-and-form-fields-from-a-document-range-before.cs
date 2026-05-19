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

            // Add some text with a bookmark.
            builder.StartBookmark("SampleBookmark");
            builder.Write("This text is inside a bookmark.");
            builder.EndBookmark("SampleBookmark");
            builder.Writeln();

            // Insert a check box form field.
            builder.InsertCheckBox("SampleCheckBox", false, 50);
            builder.Writeln();

            // Insert a combo box form field.
            builder.InsertComboBox("SampleComboBox", new[] { "Option1", "Option2", "Option3" }, 0);
            builder.Writeln();

            // At this point the document contains bookmarks and form fields.
            // Remove all bookmarks from the whole document range.
            doc.Range.Bookmarks.Clear();

            // Remove all form fields from the whole document range.
            doc.Range.FormFields.Clear();

            // Save the cleaned document.
            const string outputPath = "CleanedDocument.docx";
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Document saved to '{outputPath}' with all bookmarks and form fields removed.");
        }
    }
}
