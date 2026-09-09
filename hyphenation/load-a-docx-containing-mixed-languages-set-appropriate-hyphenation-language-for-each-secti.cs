using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Output file names.
        const string docPath = "mixed.docx";
        const string pdfPath = "hyphenated_output.pdf";
        const string enDictPath = "hyph_en_US.dic";
        const string deDictPath = "hyph_de_CH.dic";

        // -----------------------------------------------------------------
        // 1. Create minimal hyphenation dictionaries for English and German.
        // -----------------------------------------------------------------
        File.WriteAllText(enDictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        File.WriteAllText(deDictPath,
            "UTF-8\n" +
            "außergewöhnlichkeitsbegründung=au-ßer-gewön-lich-keits-be-grün-dung\n" +
            "kommunikation=ko-mmu-ni-ka-tion\n");

        // -----------------------------------------------------------------
        // 2. Register the dictionaries so Aspose.Words can hyphenate.
        // -----------------------------------------------------------------
        Hyphenation.RegisterDictionary("en-US", enDictPath);
        Hyphenation.RegisterDictionary("de-CH", deDictPath);

        // -----------------------------------------------------------------
        // 3. Build a sample document containing English and German text.
        // -----------------------------------------------------------------
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);

        // Narrow the page to force line wrapping where hyphenation can occur.
        tempDoc.FirstSection.PageSetup.PageWidth = 300; // points
        tempDoc.FirstSection.PageSetup.LeftMargin = 20;
        tempDoc.FirstSection.PageSetup.RightMargin = 20;

        // English section.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // German section.
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln("außergewöhnlichkeitsbegründung kommunikation");

        // Save the document to disk – this simulates loading an existing file later.
        tempDoc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Load the document, enable automatic hyphenation, and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);

        // Ensure the page setup is still narrow (in case the loaded doc differs).
        doc.FirstSection.PageSetup.PageWidth = 300;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.HyphenationZone = 360; // default value

        // Save the result as PDF to render hyphenation.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The expected PDF output was not created.");

        // Optional clean‑up (commented out to keep files for inspection).
        // File.Delete(docPath);
        // File.Delete(enDictPath);
        // File.Delete(deDictPath);
    }
}
