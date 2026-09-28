using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary.
        const string dictPath = "hyph_en_US.dic";
        string dictContent = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(dictPath, dictContent);

        // Register the dictionary for English (United States).
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the language for hyphenation.
        builder.Font.LocaleId = CultureInfo.GetCultureInfo("en-US").LCID;

        // Configure a narrow page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        PageSetup pageSetup = section.PageSetup;
        pageSetup.PageWidth = 200;   // points
        pageSetup.LeftMargin = 20;   // points
        pageSetup.RightMargin = 20;  // points

        // Add text containing words that have hyphenation points defined in the dictionary.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document as PDF.
        const string pdfPath = "hyphenated.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected PDF output file was not created.");
        }
    }
}
