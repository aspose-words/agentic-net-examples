using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

namespace AsposeWordsPdfRender
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (DOCX, DOC, etc.).
            string sourceDocumentPath = @"C:\Data\MyDocument.docx";

            // Path to the folder that contains the custom fonts required for rendering.
            string customFontsFolder = @"C:\Data\Fonts";

            // Path where the resulting PDF will be saved.
            string outputPdfPath = @"C:\Data\RenderedDocument.pdf";

            // Load the document using the provided Document constructor.
            Document doc = new Document(sourceDocumentPath);

            // Configure font sources so that Aspose.Words can locate fonts in the specified folder.
            // Preserve any existing font sources and add the custom folder source.
            FontSourceBase[] existingSources = FontSettings.DefaultInstance.GetFontsSources();
            FolderFontSource folderSource = new FolderFontSource(customFontsFolder, true);
            FontSettings.DefaultInstance.SetFontsSources(
                new FontSourceBase[] { existingSources[0], folderSource });

            // Create a PdfSaveOptions object to control PDF rendering.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Enable high‑quality rendering for better visual fidelity.
                UseHighQualityRendering = true,

                // Embed full fonts to ensure the PDF looks the same on any system.
                EmbedFullFonts = true
            };

            // Save the document as PDF using the provided Save method overload that accepts SaveOptions.
            doc.Save(outputPdfPath, pdfOptions);
        }
    }
}
