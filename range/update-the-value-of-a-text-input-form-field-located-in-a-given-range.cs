using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field named "MyField" with default text "Old value".
            builder.InsertTextInput("MyField", TextFormFieldType.Regular, "", "Old value", 50);

            // Save the initial document.
            const string originalPath = "Original.docx";
            doc.Save(originalPath);

            // Load the document from the saved file.
            Document loadedDoc = new Document(originalPath);

            // Locate the form field within the document's range by its name.
            FormField formField = loadedDoc.Range.FormFields["MyField"];
            if (formField != null)
            {
                // Update the value of the text input form field.
                formField.Result = "Updated value";
            }

            // Save the updated document.
            const string updatedPath = "Updated.docx";
            loadedDoc.Save(updatedPath);
        }
    }
}
