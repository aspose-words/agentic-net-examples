using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large PDF (e.g., 500 pages) as a sample source.
        const string sourcePdfPath = "large.pdf";
        CreateLargePdf(sourcePdfPath, 500);

        // Load the large PDF from a file stream.
        using (FileStream sourceStream = new FileStream(sourcePdfPath, FileMode.Open, FileAccess.Read))
        {
            Document pdfDocument = new Document(sourceStream);

            // Configure PDF/A‑2b compliance options.
            // The PdfCompliance enum version available in the referenced Aspose.Words package
            // provides PdfA2a but not PdfA2b. Using PdfA2a here ensures the code compiles;
            // it still demonstrates PDF/A‑2 compliance.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2a
            };

            // Stream the converted PDF/A‑2b (or PDF/A‑2a) directly to a MemoryStream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                pdfDocument.Save(outputStream, saveOptions);

                // Validate that data was written.
                if (outputStream.Length == 0)
                {
                    throw new InvalidOperationException("The PDF/A conversion produced an empty stream.");
                }

                // Optionally write the result to a file for verification.
                const string resultPath = "output_pdfa2b.pdf";
                File.WriteAllBytes(resultPath, outputStream.ToArray());

                if (!File.Exists(resultPath) || new FileInfo(resultPath).Length == 0)
                {
                    throw new InvalidOperationException("The PDF/A output file was not created correctly.");
                }
            }
        }
    }

    private static void CreateLargePdf(string filePath, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i} of a large PDF document.");
            if (i < pageCount)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        doc.Save(filePath, SaveFormat.Pdf);

        if (!File.Exists(filePath) || new FileInfo(filePath).Length == 0)
        {
            throw new InvalidOperationException("Failed to create the sample large PDF.");
        }
    }
}
