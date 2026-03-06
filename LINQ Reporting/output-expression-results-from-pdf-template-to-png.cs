using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template that contains expression fields.
        Document doc = new Document("Template.pdf");

        // Evaluate all fields (including expression fields) in the document.
        doc.UpdateFields();

        // Create image save options for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optionally set resolution or other rendering options.
            // Resolution = 300
        };

        // Save each page as a separate PNG file.
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Render only the current page.
            pngOptions.PageSet = new PageSet(i);
            string outputPath = $"Output_page{i + 1}.png";
            doc.Save(outputPath, pngOptions);
        }
    }
}
