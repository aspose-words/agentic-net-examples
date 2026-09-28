using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items for the combo box.
        string[] comboItems = { "Option A", "Option B", "Option C" };

        // Insert a combo box form field with a default selected index (e.g., 1 -> "Option B").
        int defaultSelectedIndex = 1;
        builder.InsertComboBox("SampleComboBox", comboItems, defaultSelectedIndex);

        // Save the document to disk.
        doc.Save("ComboBoxFormField.docx");
    }
}
