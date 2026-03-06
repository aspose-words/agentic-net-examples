using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Load the source DOCX document that contains MERGEFIELDs.
        // Replace "Template.docx" with the actual path to your template file.
        Document doc = new Document("Template.docx");

        // Define the merge field names present in the template and the values to insert.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute a simple mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as a PNG image.
        // The ImageSaveOptions allow us to control rendering; here we render the first page only.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // zero‑based index of the page to render
        };

        // Replace "Result.png" with the desired output file path.
        doc.Save("Result.png", options);
    }
}
