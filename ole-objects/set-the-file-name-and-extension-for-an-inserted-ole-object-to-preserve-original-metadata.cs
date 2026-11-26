using System;

public class Program
{
    public static void Main()
    {
        // Simulate an OLE object that would be inserted into a document.
        var oleObject = new OleObject();

        // Original file metadata.
        string originalFilePath = @"C:\Documents\Report.pdf";

        // Extract file name and extension.
        string fileName = System.IO.Path.GetFileNameWithoutExtension(originalFilePath);
        string extension = System.IO.Path.GetExtension(originalFilePath);

        // Set the metadata on the OLE object.
        oleObject.FileName = fileName;
        oleObject.Extension = extension;

        // Output the result to verify that the metadata is preserved.
        Console.WriteLine($"OLE Object inserted:");
        Console.WriteLine($"  File Name : {oleObject.FileName}");
        Console.WriteLine($"  Extension : {oleObject.Extension}");
    }
}

// Simple representation of an OLE object for demonstration purposes.
public class OleObject
{
    // The name of the original file without extension.
    public string FileName { get; set; }

    // The file extension, including the leading dot (e.g., ".pdf").
    public string Extension { get; set; }

    // Additional properties and methods that would exist in a real OLE object
    // could be added here (e.g., binary data, display options, etc.).
}
