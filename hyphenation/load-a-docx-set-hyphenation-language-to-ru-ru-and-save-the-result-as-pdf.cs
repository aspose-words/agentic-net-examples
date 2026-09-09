using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a minimal Russian hyphenation dictionary.
        string dictPath = Path.Combine(workDir, "hyph_ru_RU.dic");
        File.WriteAllText(dictPath,
@"UTF-8
привет=при-вет
автоматизация=ав-то-ма-ти-за-ция
разработка=раз-ра-бот-ка
программного=про-грамм-но-го
обеспечения=об-еспе-че-ния");

        // Register the dictionary for the ru-RU language.
        Hyphenation.RegisterDictionary("ru-RU", dictPath);

        // Create a sample DOCX with Russian text.
        string docPath = Path.Combine(workDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("ru-RU").LCID;
        builder.Writeln("Приветствую всех участников конференции, где обсуждаются вопросы автоматизации и разработки программного обеспечения.");
        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.Save(docPath);

        // Load the created document.
        Document loadedDoc = new Document(docPath);

        // Save the document as PDF.
        string pdfPath = Path.Combine(workDir, "output.pdf");
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(dictPath);
        // File.Delete(docPath);
        // File.Delete(pdfPath);
    }
}
