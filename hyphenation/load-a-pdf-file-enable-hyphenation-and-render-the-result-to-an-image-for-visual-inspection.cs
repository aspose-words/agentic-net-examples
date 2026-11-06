using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationExample
{
    public static void Main()
    {
        // Ensure a clean environment.
        const string dictionaryPath = "hyph_en_US.dic";
        const string pdfPath = "sample.pdf";
        const string imagePath = "sample.png";

        // Delete previous artifacts if they exist.
        if (File.Exists(dictionaryPath)) File.Delete(dictionaryPath);
        if (File.Exists(pdfPath)) File.Delete(pdfPath);
        if (File.Exists(imagePath)) File.Delete(imagePath);

        // 1. Create a minimal hyphenation dictionary for English (en-US).
        // The dictionary format: first line is "UTF-8", subsequent lines contain word=hyphenated-pieces.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // 2. Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // 3. Create a document with narrow page width to force line wrapping and hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the locale to English (United States) so the registered dictionary is used.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Narrow page width (points). 1 inch = 72 points.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // approx 2.78 inches
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Write a paragraph containing long words that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // 4. Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // 5. Load the PDF file back into a Document.
        Document loadedPdf = new Document(pdfPath);

        // 6. Render the first page of the loaded PDF to an image (PNG).
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // first page (zero‑based index)
        };
        loadedPdf.Save(imagePath, imageOptions);

        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Image file was not created.");

        // Example completed successfully.
        Console.WriteLine("Hyphenation PDF rendered to image successfully.");
    }
}
