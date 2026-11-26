using System;
using Aspose.Words;

namespace ComboBoxFormFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the items for the combo box.
            string[] items = { "Option A", "Option B", "Option C" };

            // Set the default selected index (0‑based). Here we select "Option B".
            int defaultSelectedIndex = 1;

            // Insert the combo box form field with a name, the items, and the default index.
            builder.InsertComboBox("SampleComboBox", items, defaultSelectedIndex);

            // Save the document to disk.
            doc.Save("ComboBoxFormField.docx");
        }
    }
}
