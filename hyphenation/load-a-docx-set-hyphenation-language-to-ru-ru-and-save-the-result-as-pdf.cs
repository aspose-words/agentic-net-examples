using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX with Russian text and save it locally.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(dataDir, "Sample.docx");
        CreateSampleDocument(sourceDocPath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX, enable automatic hyphenation and ensure the
        //    language of the text is set to Russian (ru‑RU).
        // -----------------------------------------------------------------
        Document doc = new Document(sourceDocPath);

        // Set the locale of all runs to Russian.
        int russianLcid = new CultureInfo("ru-RU").LCID;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = russianLcid;
        }

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // -----------------------------------------------------------------
        // 3. Save the document as PDF.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(dataDir, "Result.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // The example finishes without requiring any user interaction.
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a relatively large page width to make hyphenation visible.
        builder.PageSetup.PageWidth = 400; // points

        // Set Russian locale for the text.
        int russianLcid = new CultureInfo("ru-RU").LCID;
        builder.Font.LocaleId = russianLcid;
        builder.Font.Size = 24;

        // Russian text long enough to trigger hyphenation.
        string russianText = "Очень длинный русский текст, который будет перенесен на несколько строк, " +
                             "чтобы продемонстрировать работу переноса слов с помощью автоматических переносов. " +
                             "Гипенизация улучшает читаемость и внешний вид документа, особенно при узких колонках.";

        builder.Writeln(russianText);
        doc.Save(filePath);
    }
}
