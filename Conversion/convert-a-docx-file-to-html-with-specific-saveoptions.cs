using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path where the resulting HTML file will be saved.
        string outputPath = @"C:\Docs\Sample.html";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure HTML save options.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export images as separate files (not embedded as Base64).
            ExportImagesAsBase64 = false,

            // Folder where the exported images will be placed.
            ImagesFolder = Path.Combine(Path.GetDirectoryName(outputPath), "Images"),

            // Write CSS to an external stylesheet.
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = Path.Combine(Path.GetDirectoryName(outputPath), "styles.css"),

            // Produce nicely indented (pretty) HTML output.
            PrettyFormat = true
        };

        // Ensure the images folder exists before saving.
        Directory.CreateDirectory(options.ImagesFolder);

        // Save the document as HTML using the specified options.
        doc.Save(outputPath, options);
    }
}
