using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class InsertComboBoxIntoMhtml
{
    static void Main()
    {
        // Path to the source MHTML document.
        string inputPath = "input.mhtml";

        // Path where the modified MHTML document will be saved.
        string outputPath = "output.mhtml";

        // Load the MHTML document. Use HtmlLoadOptions to ensure that
        // <select> elements are imported as form fields (not as StructuredDocumentTags).
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            PreferredControlType = HtmlControlType.FormField
        };
        Document doc = new Document(inputPath, loadOptions);

        // Create a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph to host the combo box (optional, for layout).
        builder.InsertParagraph();

        // Define the items that will appear in the combo box.
        string[] comboItems = new[]
        {
            "-- Select an option --",
            "Option A",
            "Option B",
            "Option C"
        };

        // Insert the combo box form field.
        // Parameters: field name, items array, default selected index.
        builder.InsertComboBox("MyComboBox", comboItems, 0);

        // Save the modified document back to MHTML format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
