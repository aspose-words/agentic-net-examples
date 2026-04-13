using System;
using System.IO;
using System.Globalization;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple document with English locale text that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln(
            "This is a sample paragraph containing averylongwordthatmightneedhyphenation to demonstrate hyphenation handling.");

        // Turn on automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Create an intentionally invalid hyphenation dictionary file.
        string invalidDicPath = Path.Combine(outputDir, "invalid.dic");
        File.WriteAllText(invalidDicPath, "This is not a valid hyphenation dictionary format.");

        // Try to register the invalid dictionary and catch any errors.
        try
        {
            Hyphenation.RegisterDictionary("en-US", invalidDicPath);
            Console.WriteLine("Hyphenation dictionary registered successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"Failed to register hyphenation dictionary for language 'en-US'. Reason: {ex.Message}");
        }

        // Save the document to PDF and report the result.
        string pdfPath = Path.Combine(outputDir, "HyphenatedDocument.pdf");
        try
        {
            doc.Save(pdfPath);
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"Document saved successfully to '{pdfPath}'.");
            }
            else
            {
                Console.WriteLine($"Document save reported success, but file was not found at '{pdfPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving document: {ex.Message}");
        }
    }
}
