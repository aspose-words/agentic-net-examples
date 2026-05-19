using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace BatchDocxToPdfWithWatermark
{
    public class Program
    {
        // Base64‑encoded PNG (1 × 1 pixel, solid blue). This provides a deterministic image without using System.Drawing.
        private const string LogoBase64 = 
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6cVYAAAAAElFTkSuQmCC";

        private static readonly string WorkingFolder = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        private static readonly string InputFolder   = Path.Combine(WorkingFolder, "InputDocs");
        private static readonly string OutputFolder  = Path.Combine(WorkingFolder, "OutputPdfs");
        private static readonly string LogoPath      = Path.Combine(WorkingFolder, "logo.png");

        public static void Main()
        {
            PrepareFolders();
            CreateSampleLogo();
            CreateSampleDocuments();

            // Process each DOCX file: add image watermark and save as PDF.
            foreach (string docxPath in Directory.GetFiles(InputFolder, "*.docx"))
            {
                // Load the source document.
                Document doc = new Document(docxPath);

                // Configure watermark options (no washout, default scale).
                ImageWatermarkOptions wmOptions = new ImageWatermarkOptions
                {
                    IsWashout = false
                };

                // Apply the image watermark using the logo file.
                doc.Watermark.SetImage(LogoPath, wmOptions);

                // Determine output PDF path.
                string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
                string pdfPath = Path.Combine(OutputFolder, pdfFileName);

                // Save the document as PDF.
                doc.Save(pdfPath, SaveFormat.Pdf);
            }

            // Simple verification – list generated PDFs.
            Console.WriteLine("Generated PDF files:");
            foreach (string pdf in Directory.GetFiles(OutputFolder, "*.pdf"))
            {
                Console.WriteLine(pdf);
            }
        }

        // Ensure all required directories exist.
        private static void PrepareFolders()
        {
            Directory.CreateDirectory(WorkingFolder);
            Directory.CreateDirectory(InputFolder);
            Directory.CreateDirectory(OutputFolder);
        }

        // Write the embedded logo PNG to disk.
        private static void CreateSampleLogo()
        {
            byte[] logoBytes = Convert.FromBase64String(LogoBase64);
            File.WriteAllBytes(LogoPath, logoBytes);
        }

        // Generate a few simple DOCX files for the batch operation.
        private static void CreateSampleDocuments()
        {
            for (int i = 1; i <= 3; i++)
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"Sample document #{i}");
                builder.Writeln("This document will be converted to PDF with a corporate logo watermark.");
                string docxPath = Path.Combine(InputFolder, $"Sample{i}.docx");
                doc.Save(docxPath);
            }
        }
    }
}
