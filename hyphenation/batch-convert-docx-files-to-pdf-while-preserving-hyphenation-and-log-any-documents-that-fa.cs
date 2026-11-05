using System;
using System.IO;
using System.Globalization;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "input_docs";
        string outputFolder = "output_pdfs";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a minimal hyphenation dictionary for en-US
        string dictionaryPath = "hyph_en_US.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ninternationalization=in-ter-na-tion-al-i-za-tion\ncommunication=com-mu-ni-ca-tion\n");

        // Register the dictionary
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create sample DOCX files
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set narrow page width to force line wrapping and hyphenation
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Set language for hyphenation
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;

            // Write text that contains words present in the dictionary
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            doc.Save(docPath, SaveFormat.Docx);
        }

        // Log file for failures
        string logPath = "failed.txt";
        if (File.Exists(logPath))
            File.Delete(logPath);

        // Process each DOCX file
        foreach (string docxFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            try
            {
                Document document = new Document(docxFile);
                string pdfFileName = Path.GetFileNameWithoutExtension(docxFile) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);
                document.Save(pdfPath, SaveFormat.Pdf);

                // Verify PDF creation
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException($"PDF was not created for {docxFile}.");
            }
            catch (Exception ex)
            {
                // Log failure
                File.AppendAllText(logPath, $"{docxFile}: {ex.Message}{Environment.NewLine}");
            }
        }

        // Ensure the log file exists (even if empty) for consistency
        if (!File.Exists(logPath))
            File.WriteAllText(logPath, string.Empty);
    }
}
