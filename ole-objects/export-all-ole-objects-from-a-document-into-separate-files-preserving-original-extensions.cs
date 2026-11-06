using System;
using System.IO;
using System.IO.Compression;

public class OleObjectExporter
{
    public static void Main(string[] args)
    {
        // Expect the first argument to be the path to a .docx (or .xlsx) file.
        if (args.Length == 0)
        {
            // No input file provided; nothing to do.
            return;
        }

        string documentPath = args[0];

        if (!File.Exists(documentPath))
        {
            // Input file does not exist; exit silently.
            return;
        }

        // Create an output directory next to the source file.
        string outputDirectory = Path.Combine(Path.GetDirectoryName(documentPath) ?? "", "ExtractedOleObjects");
        Directory.CreateDirectory(outputDirectory);

        // Open the Office Open XML package (ZIP archive) and extract OLE object parts.
        using (ZipArchive archive = ZipFile.OpenRead(documentPath))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                // OLE objects are typically stored under the "embeddings" folder.
                // For Word documents: "word/embeddings/", for Excel: "xl/embeddings/", etc.
                if (IsOleObjectEntry(entry.FullName))
                {
                    string fileName = Path.GetFileName(entry.FullName);
                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    string destinationPath = Path.Combine(outputDirectory, fileName);
                    entry.ExtractToFile(destinationPath, overwrite: true);
                }
            }
        }
    }

    private static bool IsOleObjectEntry(string entryPath)
    {
        // Normalize to forward slashes for consistency.
        string normalized = entryPath.Replace('\\', '/').ToLowerInvariant();

        // Check common embedding folders.
        return normalized.StartsWith("word/embeddings/") ||
               normalized.StartsWith("xl/embeddings/") ||
               normalized.StartsWith("ppt/embeddings/");
    }
}
