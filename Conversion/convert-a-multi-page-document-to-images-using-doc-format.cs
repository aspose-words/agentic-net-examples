using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        const string inputFile = @"C:\Docs\MultiPageDocument.docx";
        const string outputFolder = @"C:\Docs\PageImages";

        Directory.CreateDirectory(outputFolder);

        // Load the source document.
        Document doc = new Document(inputFile);

        // Configure image save options – PNG format with 300 DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300
        };

        // Export each page of the document to a separate PNG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Restrict the save operation to the current page only.
            saveOptions.PageSet = new PageSet(pageIndex);

            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            doc.Save(outputPath, saveOptions);
        }
    }
}
