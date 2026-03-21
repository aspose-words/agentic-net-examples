using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class HeaderCopyExample
{
    static void Main()
    {
        const string srcPath = "Source.docx";
        const string dstPath = "Destination.docx";
        const string resultPath = "Result.docx";

        // Create a source document with a primary header if it does not exist.
        if (!File.Exists(srcPath))
        {
            Document srcCreate = new Document();
            Section srcSection = srcCreate.FirstSection;

            // Add a primary header.
            HeaderFooter srcHeaderCreate = new HeaderFooter(srcCreate, HeaderFooterType.HeaderPrimary);
            Paragraph headerPara = new Paragraph(srcCreate);
            Run headerRun = new Run(srcCreate, "Source Header");
            headerPara.AppendChild(headerRun);
            srcHeaderCreate.AppendChild(headerPara);
            srcSection.HeadersFooters.Add(srcHeaderCreate);

            // Add some body content.
            Paragraph bodyPara = new Paragraph(srcCreate);
            bodyPara.AppendChild(new Run(srcCreate, "This is the source document body."));
            srcSection.Body.AppendChild(bodyPara);

            srcCreate.Save(srcPath);
        }

        // Create a destination document if it does not exist.
        if (!File.Exists(dstPath))
        {
            Document dstCreate = new Document();
            Section dstSection = dstCreate.FirstSection;

            // Add some body content.
            Paragraph bodyPara = new Paragraph(dstCreate);
            bodyPara.AppendChild(new Run(dstCreate, "This is the destination document body."));
            dstSection.Body.AppendChild(bodyPara);

            dstCreate.Save(dstPath);
        }

        // Load the source and destination documents.
        Document srcDoc = new Document(srcPath);
        Document dstDoc = new Document(dstPath);

        // Retrieve the primary header from the first section of the source document.
        HeaderFooter srcHeader = srcDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];

        // Proceed only if the source header exists and contains content.
        if (srcHeader != null && srcHeader.HasChildNodes)
        {
            // Configure import options so that the source header's formatting is preserved.
            ImportFormatOptions importOptions = new ImportFormatOptions
            {
                // When set to false, the source header/footer formatting is NOT ignored.
                IgnoreHeaderFooter = false
            };

            // Create a NodeImporter that will handle the translation of styles, lists, etc.
            NodeImporter importer = new NodeImporter(
                srcDoc,               // source document
                dstDoc,               // destination document
                ImportFormatMode.KeepSourceFormatting, // keep source formatting
                importOptions);       // apply the options defined above

            // Import (deep clone) the header node into the destination document.
            HeaderFooter importedHeader = (HeaderFooter)importer.ImportNode(srcHeader, true);

            // Ensure the destination section does not already have a primary header.
            // If it does, remove it before adding the imported one.
            Section dstSection = dstDoc.FirstSection;
            HeaderFooter existingHeader = dstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            existingHeader?.Remove();

            // Add the imported header to the destination section's HeadersFooters collection.
            dstSection.HeadersFooters.Add(importedHeader);
        }

        // Save the modified destination document.
        dstDoc.Save(resultPath);
        Console.WriteLine($"Header copied successfully. Result saved to '{resultPath}'.");
    }
}
