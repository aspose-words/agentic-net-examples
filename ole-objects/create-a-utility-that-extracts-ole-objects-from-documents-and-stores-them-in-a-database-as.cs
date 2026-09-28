using System;
using System.IO;
using System.IO.Compression;
using System.Text;

public class OleExtractor
{
    public static void Main()
    {
        string docxPath = "sample.docx";
        CreateSampleDocx(docxPath);

        string dbPath = "OleObjects.db";
        var db = new SimpleBlobDatabase(dbPath);

        using (var archive = ZipFile.OpenRead(docxPath))
        {
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase))
                {
                    using (var stream = entry.Open())
                    using (var ms = new MemoryStream())
                    {
                        stream.CopyTo(ms);
                        byte[] data = ms.ToArray();

                        db.Insert(entry.Name, data);
                        Console.WriteLine($"Extracted and stored OLE object: {entry.Name} ({data.Length} bytes)");
                    }
                }
            }
        }

        long count = db.CountRecords();
        Console.WriteLine($"Total OLE objects stored in database: {count}");
    }

    private static void CreateSampleDocx(string path)
    {
        if (File.Exists(path))
            return;

        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            var contentTypes = zip.CreateEntry("[Content_Types].xml");
            using (var writer = new StreamWriter(contentTypes.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
    <Default Extension=""xml"" ContentType=""application/xml""/>
    <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>
</Types>");
            }

            var rels = zip.CreateEntry("_rels/.rels");
            using (var writer = new StreamWriter(rels.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
</Relationships>");
            }

            var document = zip.CreateEntry("word/document.xml");
            using (var writer = new StreamWriter(document.Open()))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:body>
        <w:p><w:r><w:t>Sample document with OLE object.</w:t></w:r></w:p>
    </w:body>
</w:document>");
            }

            var oleEntry = zip.CreateEntry("word/embeddings/object1.bin");
            using (var stream = oleEntry.Open())
            {
                byte[] dummyData = new byte[] { 0xDE, 0xAD, 0xBE, 0xEF };
                stream.Write(dummyData, 0, dummyData.Length);
            }
        }
    }
}

public class SimpleBlobDatabase
{
    private readonly string _filePath;
    private readonly object _lock = new object();

    public SimpleBlobDatabase(string filePath)
    {
        _filePath = filePath;
        // Ensure the file exists
        if (!File.Exists(_filePath))
        {
            using (File.Create(_filePath)) { }
        }
    }

    public void Insert(string name, byte[] data)
    {
        byte[] nameBytes = Encoding.UTF8.GetBytes(name);
        using (var fs = new FileStream(_filePath, FileMode.Append, FileAccess.Write, FileShare.None))
        using (var bw = new BinaryWriter(fs))
        {
            bw.Write(nameBytes.Length);
            bw.Write(nameBytes);
            bw.Write(data.Length);
            bw.Write(data);
        }
    }

    public long CountRecords()
    {
        long count = 0;
        lock (_lock)
        {
            using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var br = new BinaryReader(fs))
            {
                while (fs.Position < fs.Length)
                {
                    // Read name length
                    int nameLen = br.ReadInt32();
                    // Skip name bytes
                    br.BaseStream.Seek(nameLen, SeekOrigin.Current);
                    // Read data length
                    int dataLen = br.ReadInt32();
                    // Skip data bytes
                    br.BaseStream.Seek(dataLen, SeekOrigin.Current);
                    count++;
                }
            }
        }
        return count;
    }
}
