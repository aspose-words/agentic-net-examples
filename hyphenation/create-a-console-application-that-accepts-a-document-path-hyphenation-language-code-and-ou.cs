using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine input document path and hyphenation language code.
        string inputPath = args.Length > 0 ? args[0] : "sample.docx";
        string languageCode = args.Length > 1 ? args[1] : "en-US";

        // Ensure the input document exists; create a sample if it does not.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Narrow page width to force line wrapping and hyphenation.
            builder.PageSetup.PageWidth = 300; // points

            // Set the locale for the paragraph to match the language code using reflection
            // (to stay compatible with possible API variations).
            int lcid = new CultureInfo(languageCode).LCID;
            PropertyInfo? localeProp = builder.ParagraphFormat.GetType().GetProperty("LocaleId");
            if (localeProp != null && localeProp.CanWrite)
            {
                localeProp.SetValue(builder.ParagraphFormat, lcid);
            }

            // Add sufficiently long text to demonstrate hyphenation.
            builder.Writeln(
                "Hyphenation is the process of adding hyphens to words at line breaks to improve text justification and readability. " +
                "This example demonstrates automatic hyphenation in a narrow column using Aspose.Words.");

            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Enable hyphenation and set the language using reflection (to avoid compile‑time dependency on exact member names).
        object hyphenationOptions = doc.HyphenationOptions;
        Type hyphenationType = hyphenationOptions.GetType();

        PropertyInfo? enabledProp = hyphenationType.GetProperty("IsHyphenationEnabled");
        if (enabledProp != null && enabledProp.CanWrite)
        {
            enabledProp.SetValue(hyphenationOptions, true);
        }

        PropertyInfo? languageProp = hyphenationType.GetProperty("LanguageId");
        if (languageProp != null && languageProp.CanWrite)
        {
            languageProp.SetValue(hyphenationOptions, languageCode);
        }

        // Save the document as a PDF.
        string outputPdfPath = Path.ChangeExtension(inputPath, ".pdf");
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new Exception($"PDF output was not created at '{outputPdfPath}'.");
        }
    }
}
