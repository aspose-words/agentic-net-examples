using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;

public class Program
{
    public static void Main()
    {
        // Path to the .docx file (adjust as needed)
        string docxPath = "sample.docx";

        if (!File.Exists(docxPath))
        {
            // No input file; nothing to do.
            return;
        }

        // Output directory
        string outputDir = Path.Combine(Path.GetDirectoryName(docxPath) ?? "", "ExportedOleObjects");
        Directory.CreateDirectory(outputDir);

        // Mapping of ProgID to file extension
        var progIdToExt = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Excel.Sheet.8", ".xls" },
            { "Excel.Sheet.12", ".xlsx" },
            { "Word.Document.8", ".doc" },
            { "Word.Document.12", ".docx" },
            { "PowerPoint.Show.8", ".ppt" },
            { "PowerPoint.Show.12", ".pptx" },
            { "Package", ".bin" } // generic package
        };

        // Load document.xml to extract ProgIDs
        List<string> progIds = new List<string>();
        using (var zip = ZipFile.OpenRead(docxPath))
        {
            var docEntry = zip.GetEntry("word/document.xml");
            if (docEntry != null)
            {
                using (var stream = docEntry.Open())
                {
                    XDocument doc = XDocument.Load(stream);
                    XNamespace oNs = "urn:schemas-microsoft-com:office:office";
                    foreach (var oleObj in doc.Descendants(oNs + "OLEObject"))
                    {
                        var progIdAttr = oleObj.Attribute("ProgID");
                        if (progIdAttr != null)
                        {
                            progIds.Add(progIdAttr.Value);
                        }
                        else
                        {
                            progIds.Add(string.Empty);
                        }
                    }
                }
            }

            // Get all embedding entries (usually .bin files)
            var embeddingEntries = zip.Entries
                .Where(e => e.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase) && e.Name.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                .OrderBy(e => e.Name)
                .ToList();

            int count = Math.Min(embeddingEntries.Count, progIds.Count);
            for (int i = 0; i < embeddingEntries.Count; i++)
            {
                var entry = embeddingEntries[i];
                string progId = i < progIds.Count ? progIds[i] : string.Empty;
                string ext = ".bin";

                if (!string.IsNullOrEmpty(progId) && progIdToExt.TryGetValue(progId, out string mappedExt))
                {
                    ext = mappedExt;
                }

                string outputFileName = $"OleObject{i + 1}{ext}";
                string outputPath = Path.Combine(outputDir, outputFileName);

                using (var entryStream = entry.Open())
                using (var outStream = File.Create(outputPath))
                {
                    entryStream.CopyTo(outStream);
                }
            }
        }
    }
}
