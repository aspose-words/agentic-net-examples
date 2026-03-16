using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class ContractRevisionDemo
{
    static void Main()
    {
        // Paths to the three contract versions.
        const string version1Path = @"C:\Contracts\Contract_v1.docx";
        const string version2Path = @"C:\Contracts\Contract_v2.docx";
        const string version3Path = @"C:\Contracts\Contract_v3.docx";

        // Load the documents.
        Document docV1 = new Document(version1Path);
        Document docV2 = new Document(version2Path);
        Document docV3 = new Document(version3Path);

        // Ensure the documents have no revisions before comparison.
        if (docV1.Revisions.Count != 0 || docV2.Revisions.Count != 0 || docV3.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // -----------------------------------------------------------------
        // 1. Compare V1 with V2 and save the revision set.
        // -----------------------------------------------------------------
        docV1.Compare(docV2, "AuthorV1V2", DateTime.Now);
        docV1.Save(@"C:\Contracts\Revision_V1_V2.docx");

        // -----------------------------------------------------------------
        // 2. Compare V2 with V3 and save the revision set.
        // -----------------------------------------------------------------
        docV2.Compare(docV3, "AuthorV2V3", DateTime.Now);
        docV2.Save(@"C:\Contracts\Revision_V2_V3.docx");

        // -----------------------------------------------------------------
        // 3. Create a cumulative revision document:
        //    - Start with V1.
        //    - Compare V1 with V2, accept the revisions to turn V1 into V2.
        //    - Compare the updated document (now V2) with V3, leaving revisions unaccepted.
        //    - The resulting document contains revisions for both steps.
        // -----------------------------------------------------------------
        Document cumulativeDoc = new Document(version1Path);
        cumulativeDoc.Compare(docV2, "AuthorV1V2", DateTime.Now);
        // Accept the first set of revisions so the document becomes V2.
        cumulativeDoc.Revisions.AcceptAll();

        // Now compare the updated document (V2) with V3.
        cumulativeDoc.Compare(docV3, "AuthorV2V3", DateTime.Now);
        // Do NOT accept the second set of revisions – they remain as cumulative changes.
        cumulativeDoc.Save(@"C:\Contracts\Revision_Cumulative.docx");
    }
}
