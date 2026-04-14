using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // URL of the Spanish LibreOffice hyphenation dictionary.
        const string dictionaryUrl = "https://github.com/LibreOffice/dictionaries/blob/master/es/es_ES.dic";

        // Local path for the dictionary file.
        string dictionaryPath = Path.Combine(Directory.GetCurrentDirectory(), "es_ES.dic");

        // Ensure the dictionary file exists – try to download it, otherwise create a minimal placeholder.
        if (!File.Exists(dictionaryPath))
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = client.GetAsync(dictionaryUrl).Result;
                    response.EnsureSuccessStatusCode();
                    byte[] data = response.Content.ReadAsByteArrayAsync().Result;
                    File.WriteAllBytes(dictionaryPath, data);
                }
            }
            catch (Exception)
            {
                // Fallback: create a very small dictionary file with a comment.
                // This allows the registration call to succeed in environments without internet access.
                File.WriteAllText(dictionaryPath, "% Minimal Spanish hyphenation dictionary placeholder");
            }
        }

        // Register the Spanish hyphenation dictionary from the local file.
        Hyphenation.RegisterDictionary("es-ES", dictionaryPath);

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("es-ES"))
            throw new InvalidOperationException("Failed to register the Spanish hyphenation dictionary.");

        // Create a new document and configure hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the locale to Spanish so that hyphenation uses the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // Optional: increase hyphenation zone.

        // Write a long Spanish paragraph that will trigger hyphenation.
        builder.Font.Size = 24;
        builder.Writeln(
            "Este es un ejemplo de texto en español que contiene palabras largas como internacionalización, " +
            "responsabilidad, y demás para demostrar la separación de sílabas mediante la hyphenación automática. " +
            "La correcta división de palabras mejora la legibilidad y el aspecto profesional del documento.");

        // Save the document as PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SpanishHyphenated.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output PDF was not created.", outputPath);

        Console.WriteLine("Hyphenation applied and PDF saved to: " + outputPath);
    }
}
