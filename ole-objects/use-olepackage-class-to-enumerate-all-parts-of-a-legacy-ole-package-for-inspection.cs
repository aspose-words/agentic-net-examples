using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public class OlePackage
{
    private readonly Dictionary<string, MemoryStream> _data = new Dictionary<string, MemoryStream>(StringComparer.OrdinalIgnoreCase);

    public OlePackage(byte[] fileData, string fileName)
    {
        // Store the actual file contents
        _data["FileContents"] = new MemoryStream(fileData, writable: false);

        // Build a minimal FileGroupDescriptorW stream (Unicode)
        var descriptor = new MemoryStream();
        using (var writer = new BinaryWriter(descriptor, Encoding.Unicode, leaveOpen: true))
        {
            // Number of items (DWORD)
            writer.Write((uint)1);

            // 64 bytes of placeholder data (as per the original example)
            writer.Write(new byte[64]);

            // File name as 260 WCHARs (520 bytes). Pad with zeros.
            char[] nameBuffer = new char[260];
            int copyLen = Math.Min(fileName.Length, 260);
            fileName.CopyTo(0, nameBuffer, 0, copyLen);
            writer.Write(nameBuffer);
        }
        descriptor.Position = 0;
        _data["FileGroupDescriptorW"] = descriptor;
    }

    public object GetData(string format)
    {
        if (_data.TryGetValue(format, out var stream))
        {
            // Return a new MemoryStream that shares the same underlying buffer (read‑only)
            return new MemoryStream(stream.ToArray(), writable: false);
        }
        return null;
    }
}

public class OlePackageInspector
{
    public static void Main()
    {
        // Create sample data for the OLE package (a simple text file)
        byte[] fileData = Encoding.UTF8.GetBytes("Hello from OLE package!");
        string fileName = "example.txt";

        // Create an OlePackage containing the sample file
        OlePackage package = new OlePackage(fileData, fileName);

        // Retrieve the FileGroupDescriptorW which lists the parts (files) in the package
        MemoryStream descriptorStream = package.GetData("FileGroupDescriptorW") as MemoryStream;
        if (descriptorStream == null)
        {
            Console.WriteLine("Failed to obtain FileGroupDescriptorW.");
            return;
        }

        descriptorStream.Position = 0;
        using (BinaryReader reader = new BinaryReader(descriptorStream, Encoding.Unicode, leaveOpen: true))
        {
            uint itemCount = reader.ReadUInt32();
            Console.WriteLine($"Number of items in OLE package: {itemCount}");

            for (int i = 0; i < itemCount; i++)
            {
                // Skip fields we are not interested in (total 64 bytes)
                reader.BaseStream.Seek(64, SeekOrigin.Current);

                // Read the Unicode file name (260 WCHARs)
                char[] nameChars = reader.ReadChars(260);
                string itemName = new string(nameChars);
                int nullIndex = itemName.IndexOf('\0');
                if (nullIndex >= 0)
                    itemName = itemName.Substring(0, nullIndex);

                Console.WriteLine($"Item {i + 1} name: {itemName}");
            }
        }

        // Retrieve the actual file contents (FileContents)
        MemoryStream contentStream = package.GetData("FileContents") as MemoryStream;
        if (contentStream != null)
        {
            contentStream.Position = 0;
            using (StreamReader sr = new StreamReader(contentStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, leaveOpen: true))
            {
                string content = sr.ReadToEnd();
                Console.WriteLine($"Content of '{fileName}':");
                Console.WriteLine(content);
            }
        }
        else
        {
            Console.WriteLine("Failed to obtain FileContents.");
        }
    }
}
