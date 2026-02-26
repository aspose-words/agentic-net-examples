using System;
using Aspose.Words;

namespace AsposeWordsProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the protected document will be saved.
            string outputPath = @"C:\Docs\ProtectedDocument.docx";

            // Load the existing document from the file system.
            Document doc = new Document(inputPath);

            // Apply protection to the whole document.
            // Here we use ReadOnly protection; you can choose other ProtectionType values as needed.
            doc.Protect(ProtectionType.ReadOnly);

            // Save the protected document.
            doc.Save(outputPath);
        }
    }
}
