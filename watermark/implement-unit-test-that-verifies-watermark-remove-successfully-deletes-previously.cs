using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace Aspose.Words.Tests
{
    class Program
    {
        static void Main()
        {
            // Run the watermark removal verification.
            try
            {
                VerifyWatermarkRemove();
                Console.WriteLine("Watermark removal test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Watermark removal test failed: {ex.Message}");
                Environment.Exit(1);
            }
        }

        private static void VerifyWatermarkRemove()
        {
            // Create a temporary directory for test artifacts.
            string artifactsDir = Path.Combine(Path.GetTempPath(), "Aspose.Words.Tests", Guid.NewGuid().ToString());
            Directory.CreateDirectory(artifactsDir);

            // Create a new empty document.
            Document doc = new Document();

            // Add a text watermark.
            const string watermarkText = "Test Watermark";
            doc.Watermark.SetText(watermarkText);

            // Verify that the watermark was added.
            if (doc.Watermark.Type != WatermarkType.Text)
                throw new InvalidOperationException("Watermark should be of type Text after adding.");

            // Optionally save the document to ensure the watermark persists.
            string filePath = Path.Combine(artifactsDir, "WithWatermark.docx");
            doc.Save(filePath);

            // Remove the watermark.
            doc.Watermark.Remove();

            // Verify that the watermark has been removed.
            if (doc.Watermark.Type != WatermarkType.None)
                throw new InvalidOperationException("Watermark type should be None after removal.");

            // Save the document after removal to confirm no exceptions.
            string removedFilePath = Path.Combine(artifactsDir, "WithoutWatermark.docx");
            doc.Save(removedFilePath);

            // Clean up temporary files.
            Directory.Delete(artifactsDir, true);
        }
    }
}
