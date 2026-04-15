using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OleExtractor
{
    // Simple POCO to hold OLE object information for JSON serialization.
    private class OleRecord
    {
        public string FileName { get; set; }
        public string Base64Data { get; set; }
    }

    public static void Main()
    {
        // Create a new document and insert a dummy OLE package.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] dummyFile = Encoding.UTF8.GetBytes("Sample OLE content");
        using (MemoryStream ms = new MemoryStream(dummyFile))
        {
            // Insert the OLE object as an icon. ProgId "Package" denotes a generic OLE package.
            builder.InsertOleObject(ms, "Package", true, null);
        }

        // Path for the JSON file that will act as a simple "database".
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "OleObjects.json");

        // Ensure a fresh "database".
        if (File.Exists(jsonPath))
            File.Delete(jsonPath);

        var records = new List<OleRecord>();

        // Iterate over all shapes that may contain OLE objects.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat == null)
                continue; // Not an OLE object.

            // Get the raw binary data of the OLE object.
            byte[] oleData = oleFormat.GetRawData();

            // Determine a file name for the stored object.
            string fileName = oleFormat.SuggestedFileName;
            if (string.IsNullOrEmpty(fileName))
            {
                string ext = oleFormat.SuggestedExtension ?? ".bin";
                fileName = $"OleObject_{Guid.NewGuid()}{ext}";
            }

            // Add the record to the list (store data as Base64 to keep JSON text-friendly).
            records.Add(new OleRecord
            {
                FileName = fileName,
                Base64Data = Convert.ToBase64String(oleData)
            });
        }

        // Serialize the list to JSON and write it to the file.
        string json = JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, json);

        // Verify how many OLE objects were stored.
        int count = records.Count;
        Console.WriteLine($"Extracted and stored {count} OLE object(s) into the JSON file.");
    }
}
