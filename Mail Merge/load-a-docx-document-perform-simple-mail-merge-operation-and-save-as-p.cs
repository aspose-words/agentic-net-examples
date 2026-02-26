using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToPng
{
    static void Main()
    {
        // Path to the source DOCX template that contains MERGEFIELDs.
        string templatePath = @"C:\Docs\Template.docx";

        // Load the DOCX document.
        Document doc = new Document(templatePath);

        // Define the merge field names present in the template.
        string[] fieldNames = new string[] { "FullName", "Company" };

        // Define the corresponding values for a single record.
        object[] fieldValues = new object[] { "John Doe", "Acme Corp" };

        // Perform a simple mail merge for one record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Configure image save options to render the first page as PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pngOptions.PageSet = new PageSet(0); // Render page index 0 (first page).

        // Path to the output PNG file.
        string outputPath = @"C:\Docs\Result.png";

        // Save the merged document as a PNG image.
        doc.Save(outputPath, pngOptions);
    }
}
