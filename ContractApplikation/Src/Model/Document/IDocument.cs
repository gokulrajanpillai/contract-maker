using System;

namespace ContractApplikation.Src.Model.Document
{
    interface IDocument
    {
        IDocument LoadDocument(String filePath);

        void SaveDocument(IDocument document, String filePath);
    }
}
