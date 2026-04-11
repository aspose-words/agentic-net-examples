using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string workDir = Directory.GetCurrentDirectory();
        string inputPath = Path.Combine(workDir, "input.pdf");
        string outputPath = Path.Combine(workDir, "output.pdf");

        // -----------------------------------------------------------------
        // Create a simple source PDF. In a real scenario the PDF would be
        // an existing scanned document. Here we generate one to keep the
        // example self‑contained.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to PDF/A‑1a.");
        sourceDoc.Save(inputPath, SaveFormat.Pdf);

        // Load the PDF that we just created.
        Document pdfDoc = new Document(inputPath);

        // Configure PDF/A‑1a compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // -----------------------------------------------------------------
        // Enable OCR if the current Aspose.Words version supports it.
        // The properties are set via reflection to avoid compile‑time
        // errors on versions where they are absent.
        // -----------------------------------------------------------------
        PropertyInfo ocrModeProp = typeof(PdfSaveOptions).GetProperty("OcrMode");
        if (ocrModeProp != null && ocrModeProp.CanWrite)
        {
            // Expected enum values: OcrMode.OcrAndPreserveOriginal, OcrMode.OcrOnly, etc.
            object ocrModeValue = Enum.Parse(ocrModeProp.PropertyType, "OcrAndPreserveOriginal");
            ocrModeProp.SetValue(saveOptions, ocrModeValue);
        }

        PropertyInfo ocrLanguageProp = typeof(PdfSaveOptions).GetProperty("OcrLanguage");
        if (ocrLanguageProp != null && ocrLanguageProp.CanWrite)
        {
            // Expected enum values: OcrLanguage.English, OcrLanguage.French, etc.
            object ocrLangValue = Enum.Parse(ocrLanguageProp.PropertyType, "English");
            ocrLanguageProp.SetValue(saveOptions, ocrLangValue);
        }

        // Save the document as a searchable PDF/A‑1a file.
        pdfDoc.Save(outputPath, saveOptions);

        // -----------------------------------------------------------------
        // Validation: ensure the output file exists and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("The searchable PDF/A‑1a file was not created successfully.");
        }
    }
}
