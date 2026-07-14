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

        // Insert sample text with bookmarks.
        for (int i = 1; i <= 3; i++)
        {
            string bmName = $"Bookmark_{i}";
            builder.StartBookmark(bmName);
            builder.Write($"Text inside {bmName}. ");
            builder.EndBookmark(bmName);
        }

        // Insert a few form fields.
        builder.Writeln();
        FormField combo = builder.InsertComboBox("ComboBox", new[] { "One", "Two", "Three" }, 0);
        combo.CalculateOnExit = true;
        builder.Writeln();
        FormField check = builder.InsertCheckBox("CheckBox", false, 50);
        builder.Writeln();
        FormField text = builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Placeholder", 50);

        // At this point the document contains bookmarks and form fields.
        // Remove all bookmarks from the whole document range.
        doc.Range.Bookmarks.Clear();

        // Remove all form fields from the whole document range.
        doc.Range.FormFields.Clear();

        // Save the cleaned document.
        doc.Save("Output.docx");
    }
}
