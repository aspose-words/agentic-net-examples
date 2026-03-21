using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WatermarkBatchProcessor
{
    class Program
    {
        static void Main()
        {
            // Use folders relative to the executable so they always exist.
            string baseDir = AppContext.BaseDirectory;
            string sourceFolder = Path.Combine(baseDir, "Input");
            string targetFolder = Path.Combine(baseDir, "Output");

            Directory.CreateDirectory(sourceFolder);
            Directory.CreateDirectory(targetFolder);

            const string watermarkText = "Confidential";

            string[] files = Directory.GetFiles(sourceFolder, "*.docx");
            if (files.Length == 0)
            {
                Console.WriteLine($"No .docx files found in '{sourceFolder}'. Place files there and rerun.");
                return;
            }

            foreach (string sourcePath in files)
            {
                Document doc = new Document(sourcePath);

                TextWatermarkOptions options = new TextWatermarkOptions
                {
                    FontFamily = "Arial",
                    FontSize = 36,
                    Color = Color.Red,
                    Layout = WatermarkLayout.Diagonal
                };

                doc.Watermark.SetText(watermarkText, options);

                string targetPath = Path.Combine(targetFolder, Path.GetFileName(sourcePath));
                doc.Save(targetPath);
            }

            Console.WriteLine("Watermarking completed.");
        }
    }
}
