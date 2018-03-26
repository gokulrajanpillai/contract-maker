using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContractApplikation.Src.Helper
{
    public class Constants
    {
        public static readonly string CURRENCY_SYMBOL = "€";

        public static readonly string FILE_PATH_FILLER = "\\";

        public struct FileLocation
        {

            private static readonly string PROJECT              = AppDomain.CurrentDomain.BaseDirectory;

            public static readonly string DATASOURCE            = "Vertrag-DB.accdb";

            public static readonly string PROTOTYPE_CONTRACT    = OutputFilePath("PrototypeVertrag.docx");

            public static readonly string PROTOTYPE_COSTTABLE   = OutputFilePath("TabelleKosten.xlsx");

            public static string OutputFilePath(string filename)
            {
                return PROJECT + FILE_PATH_FILLER + filename;
            }
        }
    }
}
