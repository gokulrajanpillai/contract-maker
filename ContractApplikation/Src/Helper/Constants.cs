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

            private static readonly string PROJECT_DATA         = PROJECT + FILE_PATH_FILLER + "Data";

            private static readonly string OUTPUT               = PROJECT + FILE_PATH_FILLER + "Output";

            public static readonly string DATASOURCE            = "Vertrag-DB.accdb";

            public static readonly string PROTOTYPE_CONTRACT    = PROJECT + FILE_PATH_FILLER + "PrototypeVertrag.docx";

            public static readonly string PROTOTYPE_COSTTABLE   = PROJECT + FILE_PATH_FILLER + "TabelleKosten.xlsx";

            public static readonly string OUTPUT_FILE           = OUTPUT + FILE_PATH_FILLER + "Finished_Contract.docx";

            public static string OutputFilePath(string filename)
            {
                return PROJECT + FILE_PATH_FILLER + filename;
            }
        }
    }
}
