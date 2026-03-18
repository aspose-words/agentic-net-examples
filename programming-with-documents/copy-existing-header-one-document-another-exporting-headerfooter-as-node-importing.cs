using Aspose.Words;
using Aspose.Words.Saving;

class HeaderCopyExample
{
    static void Main()
    {
        // Load the source document that contains the header we want to copy.
        Document srcDoc = new Document("Source.docx");

        // Load (or create) the destination document where the header will be inserted.
        Document dstDoc = new Document("Destination.docx");

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
        dstDoc.Save("Result.docx");
    }
}
