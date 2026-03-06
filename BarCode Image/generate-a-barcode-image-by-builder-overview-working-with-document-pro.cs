using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class BarcodeDocumentDemo
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new document and insert a barcode image.
        // -----------------------------------------------------------------
        Document barcodeDoc = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(barcodeDoc);
        // Insert a barcode image (replace with actual barcode file path).
        const string barcodeImagePath = "barcode.png";
        builder.InsertImage(barcodeImagePath);

        // -----------------------------------------------------------------
        // 2. Protect (encrypt) the document with a password.
        // -----------------------------------------------------------------
        const string protectionPassword = "protectPwd";
        barcodeDoc.Protect(ProtectionType.ReadOnly, protectionPassword);

        // Save the protected document.
        const string protectedFile = "Protected.docx";
        barcodeDoc.Save(protectedFile);                           // save

        // -----------------------------------------------------------------
        // 3. Load the protected document using the password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(protectionPassword);
        Document loadedProtectedDoc = new Document(protectedFile, loadOptions); // load

        // -----------------------------------------------------------------
        // 4. Create a second document to compare against.
        // -----------------------------------------------------------------
        Document referenceDoc = new Document();                   // create
        DocumentBuilder refBuilder = new DocumentBuilder(referenceDoc);
        refBuilder.Writeln("Reference content without barcode.");
        const string referenceFile = "Reference.docx";
        referenceDoc.Save(referenceFile);                         // save

        // -----------------------------------------------------------------
        // 5. Compare the reference document with the loaded protected document.
        // -----------------------------------------------------------------
        referenceDoc.Compare(loadedProtectedDoc, "Comparer", DateTime.Now);
        const string comparisonResultFile = "ComparisonResult.docx";
        referenceDoc.Save(comparisonResultFile);                  // save

        // -----------------------------------------------------------------
        // 6. Optional: demonstrate saving the comparison result as PDF.
        // -----------------------------------------------------------------
        const string pdfResultFile = "ComparisonResult.pdf";
        referenceDoc.Save(pdfResultFile, SaveFormat.Pdf);
    }
}
