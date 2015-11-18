using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using SolidWorks.Interop.swdocumentmgr;


namespace SWGetReferences
{
    class Program
    {
        static void Main(string[] args)
        {

            // get args
            string strBasePath = args[0];
            string strOutFile = args[1];
            string strLicenseKey = args[2];

            // get SW document manager application
            SwDMApplication swDocMgr;
            SwDMClassFactory swClassFact = new SwDMClassFactory();

            try
            {
                swDocMgr = swClassFact.GetApplication(strLicenseKey);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to get an instance of the SolidWorks Document Manager API: " + ex.Message);
                return;
            }

            // get list of potential parent files
            string[] strFiles = Directory.GetFiles(strBasePath, "*.*", SearchOption.AllDirectories);

            // get dependencies from SW
            List<string[]> lstDepends = GetDependencies(strFiles, swDocMgr);

            // write to csv
            StringBuilder sb = new StringBuilder();
            foreach (string[] row in lstDepends)
            {
                sb.AppendLine("\"" + string.Join("\",\"", row) + "\"");
            }
            File.WriteAllText(strOutFile, sb.ToString());

        }

        static List<string[]> GetDependencies(string[] strParentNames, SwDMApplication swDocMgr)
        {

            SwDMDocument19 swDoc = default(SwDMDocument19);
            SwDMSearchOption swSearchOpt = default(SwDMSearchOption);

            // returns list of string arrays
            // 0: parent file name
            // 1: child file name
            List<string[]> listDepends = new List<string[]>();

            foreach (string ParentName in strParentNames)
            {
                // get doc type
                SwDmDocumentType swDocType = GetTypeFromString(ParentName);
                if (swDocType == SwDmDocumentType.swDmDocumentUnknown)
                {
                    Console.WriteLine("Skipping unknown file: " + ParentName);
                    continue;
                }

                // get the document
                SwDmDocumentOpenError nRetVal = 0;
                swDoc = (SwDMDocument19)swDocMgr.GetDocument(ParentName, swDocType, true, out nRetVal);
                if (SwDmDocumentOpenError.swDmDocumentOpenErrorNone != nRetVal)
                {
                    Console.WriteLine("Failed to open solidworks file: " + ParentName);
                    continue;
                }

                // get arrays of dependency info (one-dimensional)
                object oBrokenRefVar;
                object oIsVirtual;
                object oTimeStamp;
                swSearchOpt = swDocMgr.GetSearchOptionObject();
                // swSearchOpt.SearchFilters = 16;
                string[] varDepends = (string[])swDoc.GetAllExternalReferences4(swSearchOpt, out oBrokenRefVar, out oIsVirtual, out oTimeStamp);
                if (varDepends == null) continue;

                Boolean[] blnIsVirtual = (Boolean[])oIsVirtual;
                for (int i = 0; i < varDepends.Length; i++)
                {

                    // file name with absolute path
                    string ChildName = varDepends[i];

                    // only return non-virtual components
                    if ((bool)blnIsVirtual[i] != true)
                    {
                        string[] strDepend = new string[2] { ParentName, ChildName };
                        listDepends.Add(strDepend);
                    }

                }

                swDoc.CloseDoc();

            }

            return listDepends;

        }

        static SwDmDocumentType GetTypeFromString(string ModelPathName)
        {

            // ModelPathName = fully qualified name of file
            SwDmDocumentType nDocType = 0;

            // Determine type of SOLIDWORKS file based on file extension
            if (ModelPathName.ToLower().EndsWith("sldprt"))
            {
                nDocType = SwDmDocumentType.swDmDocumentPart;
            }
            else if (ModelPathName.ToLower().EndsWith("sldasm"))
            {
                nDocType = SwDmDocumentType.swDmDocumentAssembly;
            }
            else if (ModelPathName.ToLower().EndsWith("slddrw"))
            {
                nDocType = SwDmDocumentType.swDmDocumentDrawing;
            }
            else
            {
                // Not a SOLIDWORKS file
                nDocType = SwDmDocumentType.swDmDocumentUnknown;
            }

            return nDocType;

        }


    }
}
