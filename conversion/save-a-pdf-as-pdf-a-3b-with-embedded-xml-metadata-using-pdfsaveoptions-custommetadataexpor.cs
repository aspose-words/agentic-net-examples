using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for PDF/A-3b conversion with custom XML metadata.");

        // Save the source document as a temporary DOCX file (required by the lifecycle rule).
        const string tempDocPath = "temp.docx";
        sourceDoc.Save(tempDocPath, SaveFormat.Docx);

        // Load the temporary document.
        Document doc = new Document(tempDocPath);

        // Prepare custom XML metadata to be stored in the PDF.
        const string customXml = "<xmpmeta xmlns=\"adobe:ns:meta/\">\n" +
                                 "  <rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\n" +
                                 "    <rdf:Description rdf:about=\"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n" +
                                 "      <dc:title>Sample PDF/A-3b Document</dc:title>\n" +
                                 "      <dc:creator>Aspose.Words Example</dc:creator>\n" +
                                 "    </rdf:Description>\n" +
                                 "  </rdf:RDF>\n" +
                                 "</xmpmeta>";

        // Store the XML string as a custom document property.
        // It will be exported as XMP metadata when saving to PDF/A.
        doc.CustomDocumentProperties.Add("CustomXmlMetadata", customXml);

        // Configure PDF save options for PDF/A‑3b compliance and embed the custom XML metadata.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑3b compliance is achieved via the PdfA3u level (includes the “b” requirements).
            Compliance = PdfCompliance.PdfA3u,

            // Export custom document properties as XMP metadata.
            CustomPropertiesExport = PdfCustomPropertiesExport.Metadata
        };

        // Save the document as PDF/A‑3b with the specified options.
        const string outputPdfPath = "output_pdfa3b.pdf";
        doc.Save(outputPdfPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF/A-3b file was not created.");

        // Clean up temporary files.
        if (File.Exists(tempDocPath))
            File.Delete(tempDocPath);
    }
}
