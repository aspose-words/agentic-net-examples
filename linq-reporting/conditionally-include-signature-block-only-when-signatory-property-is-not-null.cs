using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace SignatureBlockExample
{
    class Program
    {
        // Example model containing a Signatory property.
        public class DocumentModel
        {
            public string Signatory { get; set; }
            public string SignatoryTitle { get; set; }
        }

        static void Main(string[] args)
        {
            // Prepare a model – replace with real data source as needed.
            var model = new DocumentModel
            {
                Signatory = "John Doe",          // Set to null to omit the signature block.
                SignatoryTitle = "Chief Officer"
            };

            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Conditionally insert a signature line only when Signatory is not null.
            if (!string.IsNullOrEmpty(model.Signatory))
            {
                // Configure the signature line options.
                SignatureLineOptions options = new SignatureLineOptions
                {
                    Signer = model.Signatory,
                    SignerTitle = model.SignatoryTitle,
                    ShowDate = true,
                    DefaultInstructions = true
                };

                // Insert the signature line at the current cursor position.
                builder.InsertSignatureLine(options);
            }

            // Save the document to disk.
            doc.Save("ConditionalSignatureBlock.docx");
        }
    }
}
