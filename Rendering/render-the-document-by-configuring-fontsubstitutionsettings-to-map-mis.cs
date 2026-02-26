using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

namespace FontSubstitutionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text using a font that is not expected to be present on the system.
            const string missingFontName = "MissingFont";
            builder.Font.Name = missingFontName;
            builder.Writeln("This line uses a missing font and will be substituted.");

            // Write another line using a font that exists, to show normal rendering.
            builder.Font.Name = "Arial";
            builder.Writeln("This line uses Arial and will be rendered as is.");

            // -----------------------------------------------------------------
            // Configure font substitution settings.
            // -----------------------------------------------------------------
            // Create a FontSettings object and assign it to the document.
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // 1) Table substitution: map the missing font to a list of fallback fonts.
            //    Aspose.Words will try the substitutes in order until it finds an available one.
            //    We also add "Georgia" as the first fallback to demonstrate explicit replacement.
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
                missingFontName, new[] { "Georgia", "Times New Roman", "Courier New" });

            // 2) Default font substitution: if the table rule cannot resolve the font,
            //    use a default font for any remaining missing fonts.
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Calibri";

            // -----------------------------------------------------------------
            // Save the document as PDF.
            // -----------------------------------------------------------------
            const string outputPath = "FontSubstitutionResult.pdf";
            doc.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
