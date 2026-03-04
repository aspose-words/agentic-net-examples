using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Load the document using the Document(string) constructor (load rule).
        Document doc = new Document(inputPath);

        // Protect the entire document with read‑only protection (protect rule).
        doc.Protect(ProtectionType.ReadOnly);

        // Save the protected document (save rule).
        string outputPath = @"C:\Docs\output_protected.docx";
        doc.Save(outputPath);
    }
}
