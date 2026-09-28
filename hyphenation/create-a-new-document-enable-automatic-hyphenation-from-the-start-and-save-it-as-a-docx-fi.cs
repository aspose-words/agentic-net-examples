using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set a narrow page width to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add sample text containing long words that can be hyphenated.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document as DOCX.
        const string outputPath = "HyphenatedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected DOCX file was not created.");
    }
}
