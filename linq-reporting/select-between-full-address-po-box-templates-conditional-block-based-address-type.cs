using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ConditionalAddressExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Insert an IF field that chooses between a PO‑Box template
        // and a full‑address template based on the value of a merge
        // field called "AddressType".
        // ------------------------------------------------------------
        // The field code is built as a plain string because the
        // IF field does not expose strongly‑typed properties in the
        // Aspose.Words API.
        // Nested ADDRESSBLOCK fields are placed inside the true/false
        // parts of the IF field.
        string ifFieldCode =
            " IF { MERGEFIELD AddressType } = \"POBox\" " +                     // condition
            "\"{ ADDRESSBLOCK \\f \\\"<Address Line 1> PO Box <Postcode> <Country>\\\" }\" " + // true part
            "\"{ ADDRESSBLOCK \\f \\\"<Address Line 1> <Region> <Postcode> <Country>\\\" }\"";   // false part

        // Insert the IF field. The second argument (fieldValue) is an empty
        // placeholder because we will let Word calculate the result later.
        builder.InsertField(ifFieldCode, "");

        // ------------------------------------------------------------
        // Update all fields so that the IF field evaluates and the
        // appropriate ADDRESSBLOCK field result is shown.
        // ------------------------------------------------------------
        doc.UpdateFields();

        // Save the document.
        doc.Save("ConditionalAddress.docx");
    }
}
