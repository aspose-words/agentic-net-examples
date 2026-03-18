using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class ExtractEmbeddedFonts
{
    static void Main()
    {
        // Directories for input and output files.
        string dataDir = @"C:\Data\";
        string outputDir = @"C:\Output\";
        Directory.CreateDirectory(outputDir);

        // Load a source document that uses TrueType fonts.
        Document srcDoc = new Document(dataDir + "Source.docx");

        // Configure the document to embed TrueType fonts when saved.
        FontInfoCollection srcFonts = srcDoc.FontInfos;
        srcFonts.EmbedTrueTypeFonts = true;
        srcFonts.EmbedSystemFonts = true; // optional, embeds system fonts as well.

        // Save the document to PDF with subsetting (default behavior).
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // creates a PdfSaveOptions instance
        pdfOptions.EmbedFullFonts = false; // false => fonts will be subsetted.
        string pdfPath = Path.Combine(outputDir, "Rendered.pdf");
        srcDoc.Save(pdfPath, pdfOptions); // saves the document as PDF.

        // Load the generated PDF.
        Document pdfDoc = new Document(pdfPath);

        // Iterate over all fonts in the PDF and extract embedded TrueType fonts.
        FontInfoCollection pdfFonts = pdfDoc.FontInfos;
        for (int i = 0; i < pdfFonts.Count; i++)
        {
            FontInfo fontInfo = pdfFonts[i];

            // Skip non‑TrueType fonts.
            if (!fontInfo.IsTrueType)
                continue;

            // Attempt to get the embedded font in OpenType format.
            byte[] fontData = fontInfo.GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular);

            // If the OpenType format is not available, try EmbeddedOpenType and convert it.
            if (fontData == null)
                fontData = fontInfo.GetEmbeddedFontAsOpenType(EmbeddedFontStyle.Regular);

            // If we obtained font bytes, write them to a .ttf file.
            if (fontData != null)
            {
                string safeName = fontInfo.Name.Replace(' ', '_');
                string fontPath = Path.Combine(outputDir, safeName + ".ttf");
                File.WriteAllBytes(fontPath, fontData);
                Console.WriteLine($"Extracted font: {fontInfo.Name} -> {fontPath}");
            }
        }
    }
}
