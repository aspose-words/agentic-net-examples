using System;
using System.IO;
using System.Text;

public class OlePackage
{
    public byte[] Data { get; private set; }
    public string Label { get; set; }

    public OlePackage(byte[] data, string label)
    {
        Data = data ?? Array.Empty<byte>();
        Label = label ?? string.Empty;
    }

    // Save the package to a file using a simple binary format:
    // [labelLength][labelUtf8][dataLength][data]
    public void SaveToFile(string path)
    {
        using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: false))
        {
            byte[] labelBytes = Encoding.UTF8.GetBytes(Label);
            writer.Write(labelBytes.Length);
            writer.Write(labelBytes);
            writer.Write(Data.Length);
            writer.Write(Data);
        }
    }

    // Load a package from a file written by SaveToFile.
    public static OlePackage LoadFromFile(string path)
    {
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: false))
        {
            int labelLength = reader.ReadInt32();
            string label = Encoding.UTF8.GetString(reader.ReadBytes(labelLength));
            int dataLength = reader.ReadInt32();
            byte[] data = reader.ReadBytes(dataLength);
            return new OlePackage(data, label);
        }
    }
}

public class OlePackageDemo
{
    public static void Main()
    {
        // Prepare temporary file path
        string tempFile = Path.Combine(Path.GetTempPath(), "OlePackageDemo.bin");

        try
        {
            // Create sample data for the OLE package
            byte[] data = Encoding.UTF8.GetBytes("Sample OLE package content");

            // Create an OlePackage with an initial label
            OlePackage package = new OlePackage(data, "Original Package");
            package.SaveToFile(tempFile);

            // Load the package from the file and display its label
            OlePackage loaded = OlePackage.LoadFromFile(tempFile);
            Console.WriteLine($"Loaded label: {loaded.Label}");

            // Modify the label and save again
            loaded.Label = "Modified Package";
            loaded.SaveToFile(tempFile);

            // Reload to verify the change
            OlePackage reloaded = OlePackage.LoadFromFile(tempFile);
            Console.WriteLine($"Reloaded label: {reloaded.Label}");
        }
        finally
        {
            // Clean up the temporary file
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }
}
