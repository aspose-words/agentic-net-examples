using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Load the source DOCX file that contains MERGEFIELDs.
        Document doc = new Document("Template.docx");

        // Define the names of the merge fields present in the template
        // and the corresponding values to insert.
        string[] fieldNames = { "FullName", "Address" };
        object[] fieldValues = { "John Doe", "123 Main St., Anytown" };

        // Perform a simple mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Configure image save options to render the document as PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Render only the first page (index 0) of the document.
        pngOptions.PageSet = new PageSet(0);

        // Save the merged document as a PNG image.
        doc.Save("MergedResult.png", pngOptions);
    }
}
