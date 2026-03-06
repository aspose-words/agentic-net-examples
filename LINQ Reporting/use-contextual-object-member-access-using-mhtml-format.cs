using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsMhtmlExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (any format supported by Aspose.Words).
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Load the source document.
            Document doc = new Document(sourcePath);

            // Save the document as MHTML using CID URLs for embedded resources.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true, // Use CID URLs.
                PrettyFormat = true                     // Optional: make the output more readable.
            };

            string mhtmlPath = @"C:\Docs\OutputDocument.mht";
            doc.Save(mhtmlPath, saveOptions);

            // Load the MHTML document back into a Document object.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Document loadedDoc = new Document(mhtmlPath, loadOptions);

            // Example of contextual object member access:
            // Access the frameset hierarchy and print information about each frame.
            if (loadedDoc.Frameset != null && loadedDoc.Frameset.ChildFramesets.Count > 0)
            {
                foreach (var framePage in loadedDoc.Frameset.ChildFramesets)
                {
                    // If this frameset contains child frames, iterate them.
                    if (framePage.ChildFramesets.Count > 0)
                    {
                        foreach (var frame in framePage.ChildFramesets)
                        {
                            Console.WriteLine($"Frame URL: {frame.FrameDefaultUrl}");
                            Console.WriteLine($"Is linked to a file: {frame.IsFrameLinkToFile}");
                        }
                    }
                    else
                    {
                        // Single frame without children.
                        Console.WriteLine($"Frame URL: {framePage.FrameDefaultUrl}");
                        Console.WriteLine($"Is linked to a file: {framePage.IsFrameLinkToFile}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No frameset information found in the loaded MHTML document.");
            }
        }
    }
}
