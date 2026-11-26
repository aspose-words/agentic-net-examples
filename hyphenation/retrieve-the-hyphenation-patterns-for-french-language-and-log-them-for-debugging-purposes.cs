using System;
using System.IO;
using System.Globalization;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal French hyphenation dictionary.
        const string dictionaryPath = "hyph_fr_FR.dic";
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinaire=ex-tra-or-di-nai-re\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Retrieve and log the hyphenation patterns.
        Console.WriteLine("Hyphenation patterns for French (fr-FR):");
        foreach (string line in File.ReadAllLines(dictionaryPath))
        {
            Console.WriteLine(line);
        }

        // Create a document to demonstrate hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = (int)CultureInfo.GetCultureInfo("fr-FR").LCID;
        builder.Writeln("extraordinaire communication extraordinaire communication");

        // Narrow page width to force hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the document as PDF.
        const string outputPath = "HyphenatedFrench.pdf";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output was not created.");
    }
}
