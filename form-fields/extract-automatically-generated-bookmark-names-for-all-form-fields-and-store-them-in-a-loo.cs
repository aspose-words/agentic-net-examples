using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        FormField textField = builder.InsertTextInput(
            "TextField1",               // name
            TextFormFieldType.Regular, // type
            "",                         // format
            "Default text",             // default value
            0);                         // max length

        // Insert a checkbox form field.
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox1", // name
            false,       // default state
            0);          // size

        // Insert a combo box (dropdown) form field.
        // The selectedIndex must be within the bounds of the items array.
        // Start with a single placeholder item.
        FormField comboBox = builder.InsertComboBox(
            "ComboBox1",                     // name
            new[] { "Placeholder" },        // initial items
            0);                              // selected index (valid)

        // Add real items to the combo box.
        comboBox.DropDownItems.Clear(); // Remove placeholder.
        comboBox.DropDownItems.Add("Item 1");
        comboBox.DropDownItems.Add("Item 2");
        comboBox.DropDownItems.Add("Item 3");

        // Save the document with the form fields.
        string filePath = "FormFields.docx";
        doc.Save(filePath);

        // Extract automatically generated bookmark names for all form fields.
        Dictionary<string, string> fieldBookmarkLookup = new Dictionary<string, string>();

        foreach (FormField field in doc.Range.FormFields)
        {
            if (field == null)
                continue;

            // Each legacy form field creates a bookmark with the same name.
            Bookmark bookmark = doc.Range.Bookmarks[field.Name];
            string bookmarkName = bookmark != null ? bookmark.Name : string.Empty;

            fieldBookmarkLookup[field.Name] = bookmarkName;
        }

        // Output the lookup dictionary (demonstration purpose).
        foreach (KeyValuePair<string, string> kvp in fieldBookmarkLookup)
        {
            Console.WriteLine($"Form Field Name: {kvp.Key}, Bookmark Name: {kvp.Value}");
        }

        // Save the processed document.
        doc.Save("FormFields_Processed.docx");
    }
}
