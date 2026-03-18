using System;
using Aspose.Words;
using Aspose.Words.Fields;

class EscapeKeywordExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an IF field where the left expression uses the reserved keyword "new".
        // In field codes the keyword must be escaped with the @ prefix: @new
        // The field will evaluate: IF @new = 5 "True" "False"
        // Since @new is not a defined variable, the field will treat it as a literal string.
        string fieldCode = " IF @new = 5 \"True\" \"False\" ";
        builder.InsertField(fieldCode, null);

        // Update fields to calculate the result.
        doc.UpdateFields();

        // Save the document.
        doc.Save("EscapedKeyword.docx");
    }
}
