using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using Codaxy.Xlio.Model.Opc;
using System.Xml.Serialization;
using Codaxy.Xlio.Model.Oxml;
using System.Diagnostics;

namespace Codaxy.Xlio.IO
{
    //SRP: Read xlsx file into Workbook
    public partial class XlsxFileReader : IDisposable
    {
        ZipFile archive;

        ContentTypes contentTypes;
        Workbook workbook;

        public XlsxFileReader(Stream input)
        {
            archive = new ZipFile(input);
        }

        #region ContentTypes

        void ReadContentTypes()
        {
            contentTypes = new ContentTypes();

            //Debug.WriteLine("");
            //Debug.WriteLine("ContentTypes:");

            var ctypes = ReadFile<CT_Types>("[Content_Types].xml");
            foreach (var item in ctypes.Items)
            {
                CT_Default ctdef;
                CT_Override ctovr;
                if (Util.TryCast<CT_Default>(item, out ctdef))
                {
                    //Debug.WriteLine(String.Format("DCT: *.{1}: {0}", ctdef.ContentType, ctdef.Extension));
                    contentTypes.RegisterDefaultType(ctdef.Extension, ctdef.ContentType);
                }
                else if (Util.TryCast<CT_Override>(item, out ctovr))
                {
                    //Debug.WriteLine(String.Format("OCT: {1}: {0}", ctovr.ContentType, ctovr.PartName));
                    contentTypes.RegisterPartTypeOverride(ctovr.PartName, ctovr.ContentType);                    
                }
            }
        }

        #endregion

        public Workbook Read()
        {
            workbook = new Workbook();

            //foreach (ZipEntry e in archive)
            //    Debug.WriteLine(String.Format("File: {0}", e.Name));

            ReadContentTypes();

            //Debug.WriteLine("");
            //Debug.WriteLine("Relationships:");
            var packageRels = ReadFile<CT_Relationships>("_rels/.rels");
            foreach (var pr in packageRels.Relationship)
            {
                //Debug.WriteLine(String.Format("{0} {1} {2}", pr.Id, pr.Target, pr.Type));
                var contentType = contentTypes.ResolvePartType(pr.Target);
                switch (contentType)
                {
                    case ContentTypes.Workbook:
                        LoadWorkbook(pr.Target, pr.Type);
                        break;
                }
            }

            return workbook;
        }

       

        private void LoadWorkbook(String path, String relType)
        {
            //Process workbook relationships

            Dictionary<String, CT_Relationship> sheetRels = new Dictionary<string, CT_Relationship>();
            String filePath;
            String fileName;
            PathUtil.SplitFilePath(path, out filePath, out fileName);
            var relsPath = PathUtil.CombinePaths(filePath, String.Format("/_rels/{0}.rels", fileName));
            if (FileExists(relsPath))
            {
                //Debug.WriteLine("");
                //Debug.WriteLine("Workbook relationships:");

                var workbookRels = ReadFile<CT_Relationships>(relsPath);
                sharedStrings = new SharedStrings();

                foreach (var rel in workbookRels.Relationship)
                {
                    //Debug.WriteLine(String.Format("{0} {1} {2}", rel.Id, rel.Target, rel.Type));
                    var partPath = PathUtil.CombinePaths(filePath, rel.Target);
                    var contentType = contentTypes.ResolvePartType(partPath);
                    switch (contentType)
                    {
                        case ContentTypes.SharedStrings:
                            ReadSharedStrings(partPath);
                            break;
                        case ContentTypes.Sheet:
                            sheetRels.Add(rel.Id, rel);
                            break;
                        case ContentTypes.Styles:
                            LoadStyles(partPath);
                            break;
                    }
                }
            }

            //Process workbook data
            var wb = ReadFile<CT_Workbook>(path);

            //Debug.WriteLine("");
            //Debug.WriteLine("Sheets: ");           
            foreach (var sheetInfo in wb.sheets)
            {
                //Debug.WriteLine(String.Format("{2} {0}: {1}", sheetInfo.name, sheetInfo.sheetId, sheetInfo.id));
                var sheet = new Sheet(sheetInfo.name);               
                CT_Relationship rel;
                if (sheetRels.TryGetValue(sheetInfo.id, out rel))
                {
                    var partPath = PathUtil.CombinePaths(filePath, rel.Target);
                    ReadSheet(partPath, sheet);
                }
                
                workbook.Sheets.AddSheet(sheet);
            }

            if (wb.definedNames != null)
            {
                foreach (var dn in wb.definedNames)
                {
                    DefinedName definedName = new DefinedName
                    {
                        Name = dn.name,
                        Value = dn.Value
                    };

                    if (dn.localSheetIdSpecified)
                        workbook.Sheets[(int)dn.localSheetId].DefinedNames.AddDefinedName(definedName);                    
                    else
                        workbook.DefinedNames.AddDefinedName(definedName);
                }
            }

        }

        SharedStrings sharedStrings;
        private void ReadSharedStrings(string path)
        {
            var sst = ReadFile<CT_Sst>(path);
            if (sst.si != null)
                foreach (var ss in sst.si)
                {
                    String s;
                    if (ss.t != null)
                        s = ss.t;
                    else if (ss.rPh != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var t in ss.rPh)
                            sb.Append(t.t);
                        s = sb.ToString();
                    }
                    else if (ss.r != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var t in ss.r)
                            sb.Append(t.t);
                        s = sb.ToString();
                    }
                    else
                        s = String.Empty;
                    sharedStrings.Add(s);
                }
        }

        #region Zip

        private string GetZipPath(string path)
        {
            return path.TrimStart('/');
        }

        bool FileExists(string filePath)
        {
            var zipPath = GetZipPath(filePath);
            var zipEntry = archive.FindEntry(zipPath, true);            
            return zipEntry != -1;
        }

        T ReadFile<T>(String fileName)
        {
            var zipPath = GetZipPath(fileName);
            var zipEntry = archive.GetEntry(zipPath);
            if (zipEntry == null)
                throw ExceptionFactory.InvalidOperation("File '{0}' not found in archive!", fileName);
            using (var stream = archive.GetInputStream(zipEntry))
			using (var reader = new StreamReader(stream, true))
            {
				XmlSerializer s = new XmlSerializer(typeof(T));
                return (T)s.Deserialize(reader);
            }
        }

        #endregion

        public void Dispose()
        {
            ((IDisposable)archive).Dispose();
        }
    }
}
