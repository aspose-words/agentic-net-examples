using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeRegionToPng
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string sourcePath = "Template.docx";

        // Load the existing document.
        Document doc = new Document(sourcePath);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a mail‑merge region named "MyRegion".
        // Insert the region start field.
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");

        // Insert a sample merge field that will be repeated inside the region.
        builder.InsertField(" MERGEFIELD Name");

        // Insert the region end field.
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Prepare image save options for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render the first page only; change PageSet if you need a different page.
            PageSet = new PageSet(0),

            // Optional: set resolution (dpi) for higher quality.
            Resolution = 300
        };

        // Save the document as a PNG image.
        const string outputPath = "Result.png";
        doc.Save(outputPath, pngOptions);
    }
}
