using System;
using System.IO;
using System.IO.Compression;

public class OleExtractor
{
    public static void Main(string[] args)
    {
        // Expect two arguments: path to the .docx file and output directory
        if (args.Length < 2)
        {
            // No interactive prompts; just exit if arguments are missing
            return;
        }

        string docxPath = args[0];
        string outputDirectory = args[1];

        if (!File.Exists(docxPath))
        {
            return;
        }

        Directory.CreateDirectory(outputDirectory);

        try
        {
            using FileStream fileStream = File.OpenRead(docxPath);
            using ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Read, leaveOpen: false);

            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                // OLE objects are stored under word/embeddings/
                if (entry.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase) &&
                    !entry.FullName.EndsWith("/", StringComparison.Ordinal))
                {
                    string fileName = Path.GetFileName(entry.FullName);
                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    string destinationPath = Path.Combine(outputDirectory, fileName);

                    using Stream entryStream = entry.Open();
                    using FileStream outputStream = File.Create(destinationPath);
                    entryStream.CopyTo(outputStream);
                }
            }
        }
        catch
        {
            // Suppress any exceptions; the program must finish on its own
        }
    }
}
