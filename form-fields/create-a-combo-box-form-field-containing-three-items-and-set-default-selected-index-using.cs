using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the items that will appear in the combo box.
        string[] comboItems = new string[] { "Option A", "Option B", "Option C" };

        // Insert a combo box form field named "SampleComboBox" with the items above.
        // The third parameter sets the default selected index (0‑based). Here we select "Option B".
        builder.InsertComboBox("SampleComboBox", comboItems, 1);

        // Save the document to disk.
        doc.Save("ComboBoxFormField.docx");
    }
}
