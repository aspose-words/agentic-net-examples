using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsTiffPageSaving
{
    public class Program
    {
        public static void Main()
        {
            // Define an output folder for the generated TIFF pages.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with three pages.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 3.");

            // Configure ImageSaveOptions for TIFF output and assign a custom callback.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
            saveOptions.PageSavingCallback = new CustomPageSavingCallback(outputDir);

            // Save the document. The callback will be invoked for each page.
            string dummyFileName = Path.Combine(outputDir, "dummy.tiff"); // Name is ignored by the callback.
            doc.Save(dummyFileName, saveOptions);

            // Verify that a TIFF file was created for each page.
            for (int i = 0; i < doc.PageCount; i++)
            {
                string expectedPath = Path.Combine(outputDir, $"Page_{i + 1}.tiff");
                if (!File.Exists(expectedPath))
                    throw new InvalidOperationException($"Expected TIFF page not found: {expectedPath}");
            }

            Console.WriteLine("All TIFF pages were saved successfully.");
        }

        // Callback that customizes the file name for each saved page.
        private class CustomPageSavingCallback : IPageSavingCallback
        {
            private readonly string _outputDirectory;

            public CustomPageSavingCallback(string outputDirectory)
            {
                _outputDirectory = outputDirectory;
            }

            public void PageSaving(PageSavingArgs args)
            {
                // PageIndex is zero‑based; add 1 for a more natural numbering.
                string fileName = Path.Combine(_outputDirectory, $"Page_{args.PageIndex + 1}.tiff");
                args.PageFileName = fileName;
                // Keep the default behavior of closing the stream after saving.
                args.KeepPageStreamOpen = false;
            }
        }
    }
}
