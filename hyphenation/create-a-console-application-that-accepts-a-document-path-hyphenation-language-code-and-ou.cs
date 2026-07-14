using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationExample
{
    public static void Main(string[] args)
    {
        // Determine input document path and language code.
        string docPath = args.Length > 0 ? args[0] : "sample.docx";
        string languageCode = args.Length > 1 ? args[1] : "en-US";

        // Prepare a minimal hyphenation dictionary for the requested language.
        string dictFileName = $"hyph_{languageCode.Replace("-", "_")}.dic";
        if (!File.Exists(dictFileName))
        {
            // The first line must specify the encoding.
            // Subsequent lines contain word=hyphenated-pieces.
            string dictContent =
                "UTF-8\n" +
                "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
                "internationalization=in-ter-na-tion-al-i-za-tion\n" +
                "communication=com-mu-ni-ca-tion\n" +
                "demonstration=de-mon-stra-tion\n" +
                "hyphenation=hy-phen-a-tion\n";

            File.WriteAllText(dictFileName, dictContent);
        }

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary(languageCode, dictFileName);

        Document doc;

        if (File.Exists(docPath))
        {
            // Load the existing document.
            doc = new Document(docPath);
        }
        else
        {
            // Create a new document with sample text that will trigger hyphenation.
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Narrow page width to force line wrapping.
            Section section = doc.FirstSection;
            section.PageSetup.PageWidth = 200; // points
            section.PageSetup.LeftMargin = 20;
            section.PageSetup.RightMargin = 20;

            // Set the document language.
            builder.Font.LocaleId = new CultureInfo(languageCode).LCID;

            // Write sample text containing words from the dictionary.
            builder.Writeln("extraordinarycharacteristically internationalization communication demonstration hyphenation");
        }

        // Ensure the document language matches the requested language.
        if (doc.Styles["Normal"]?.Font != null)
        {
            doc.Styles["Normal"].Font.LocaleId = new CultureInfo(languageCode).LCID;
        }

        // Save the document as PDF.
        string pdfPath = Path.ChangeExtension(docPath, ".pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"PDF output was not created at '{pdfPath}'.");
        }
    }
}
