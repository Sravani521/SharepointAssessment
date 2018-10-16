using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Security;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Configuration;
using System.Net.Mail;
using Microsoft.SharePoint.Client.Utilities;

namespace UploadingFiles
{
    class Program
    {

        private static SecureString GetPassword()
        {
            ConsoleKeyInfo ckinfo;
            
            SecureString securePassword = new SecureString();
            do
            {
                ckinfo = Console.ReadKey(true);
                if (ckinfo.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(ckinfo.KeyChar);
                    Console.Write(ckinfo.KeyChar);
                }
            }
            while (ckinfo.Key != ConsoleKey.Enter);
            return securePassword;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("enter username");
            string username = Console.ReadLine();
            Console.WriteLine("enter the password");
            SecureString password = GetPassword();


            using (var clientContextobj = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ForAssessment/"))
            {


                clientContextobj.Credentials = new SharePointOnlineCredentials(username, password);
                SP.List spListobj = clientContextobj.Web.Lists.GetByTitle("Documents");

                ReadFileName(clientContextobj);
                Console.WriteLine("retrieved");
            }
            Console.ReadLine();
        }
        private static void ReadFileName(ClientContext clientContextobj)
        {
            string FileName = string.Empty;

            const string Title = "Title";
            const string lstDocName = "Documents";
            const string strFolderServerRelativeUrl = "/teams/ForAssessment/Shared%20Documents";
            string strErrorMsg = string.Empty;
            try
            {
                List listobj = clientContextobj.Web.Lists.GetByTitle(lstDocName);
                Web web = clientContextobj.Web;
                clientContextobj.Load(listobj);
                clientContextobj.ExecuteQuery();

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
                //camlQuery.FolderServerRelativeUrl = strFolderServerRelativeUrl;
                camlQuery.FolderServerRelativeUrl= strFolderServerRelativeUrl + "/"+"xlsheet.xlsx";

                ListItemCollection licollectionobj = listobj.GetItems(CamlQuery.CreateAllItemsQuery());

                clientContextobj.Load(licollectionobj, items => items.Include(i => i[Title]));
                clientContextobj.ExecuteQuery();
                for (int i = 0; i < licollectionobj.Count; i++)
                {
                    SP.ListItem liobj = licollectionobj[i];
                    if (liobj[Title] != null)
                    {
                        FileName = liobj[Title].ToString();
                        if (i == 0)
                        {

                            ReadExcelData(clientContextobj, FileName);
                        }
                    }
                }

            }
            catch (Exception e)
            {

                Console.WriteLine(e.Message);

            }

            Console.ReadKey();

        }

        private static void ReadExcelData(ClientContext clientContextobj, string FileName)
        {

            const string lstDocName = "Documents";
            try
            {

                DataTable datatable = new DataTable("TempExcelDataTable");
                List list = clientContextobj.Web.Lists.GetByTitle(lstDocName);
                clientContextobj.Load(list.RootFolder);
                clientContextobj.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + "xlsheet.xlsx";
                SP.File fileobj = clientContextobj.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> clientresult = fileobj.OpenBinaryStream();
                clientContextobj.Load(fileobj);
                clientContextobj.ExecuteQuery();
                using (System.IO.MemoryStream mstream = new System.IO.MemoryStream())
                {
                    if (clientresult != null)
                    {
                        clientresult.Value.CopyTo(mstream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mstream, false))
                        {
                            WorkbookPart WBPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart WBPart1 = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = WBPart1.Worksheet;
                            SheetData sheetdata = worksheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetdata.Descendants<Row>();
                            foreach (Cell cellvalue in rows.ElementAt(0))
                            {
                                string str = GetCellValue( document, cellvalue);
                                datatable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow datarow = datatable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        datarow[i] = GetCellValue( document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    datatable.Rows.Add(datarow);
                                }
                            }
                            datatable.Rows.RemoveAt(0);
                        }
                    }
                }
                ReadData(datatable);
            }

            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
            Console.ReadKey();

        }

        public static void ReadData(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                foreach (var values in dr.ItemArray)
                {
                    Console.WriteLine(values);
                }
            }

        }
        private static void UpdateSPList(ClientContext clientContext, DataTable datatable, string filename)
        {

            ClientContext ctx = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ForAssessment/");

            List targetList = ctx.Web.Lists.GetByTitle("Documents");
            ctx.ExecuteQuery();
            string filepath= @"C:\Users\sravani.makthala\Documents\notepad files";
            FileCreationInformation fci = new FileCreationInformation();
            fci.Content = System.IO.File.ReadAllBytes(filepath);
           
            fci.Url = "SampleFile";
            fci.Overwrite = true;
            SP.File fileToUpload = targetList.RootFolder.Files.Add(fci);
            ctx.Load(fileToUpload);
            ctx.ExecuteQuery();
            //bool isError = true;
            //string strErrorMsg = string.Empty;
            //Int32 count = 0;
            //const string lstName = "xlsheet";

            //const string lstCol1 = "Created By";
            //const string lstCol2 = "typeof";
            //const string lstCol3 = "Size";

            //try
            //{
            //    string fileExtension = ".xlsx";
            //    string fileNameWithOutExtension = filename.Substring(0, filename.Length - fileExtension.Length);
            //    if (fileNameWithOutExtension.Trim() == lstName)
            //    {
            //        SP.List listobj = clientContext.Web.Lists.GetByTitle(lstName);
            //        foreach (DataRow row in datatable.Rows)
            //        {

            //            if (count == 0)
            //            {
            //                FileCreationInformation filecreationobj = new FileCreationInformation();
            //                clientContext.ExecuteQuery();
            //            }
            //            isError = false;
            //        }
            //    }
            //}
            //catch (Exception e)
            //{
            //    isError = true;
            //    strErrorMsg = e.Message;
            //}
            //finally
            //{
            //    if (isError)
            //    {
            //        // Logging;
            //    }
            //}
        }
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    // Logging
                }
            }
            return value;
        }


    }
}

//string fileNameonly = Fileupload1.FileName;  // Only file name.

//private byte[] ToByteArray(Stream inputStream)
//{
//    using (MemoryStream ms = new MemoryStream())
//    {

//        inputStream.CopyTo(ms);
//        return ms.ToArray();
//    }

//}

//private void AddFileToDocumentLibrary(string documentLibraryUrl, string filename, string Title)
//{
//    SPSecurity.RunWithElevatedPrivileges(delegate ()
//    {
//        using (SPSite site = new SPSite(documentLibraryUrl))
//        {
//            using (SPWeb web = site.OpenWeb())
//            {
//                Stream StreamImage = null;
//                if (Fileupload1.HasFile)
//                {
//                    StreamImage = Fileupload1.PostedFile.InputStream;
//                }
//                byte[] file_bytes = ToByteArray(StreamImage);
//                web.AllowUnsafeUpdates = true;
//                SPDocumentLibrary documentLibrary = (SPDocumentLibrary)web.Lists["DocumentLibraryName"];
//                SPFileCollection files = documentLibrary.RootFolder.Files;
//                SPFile newFile = files.Add(documentLibrary.RootFolder.Url + "/" + filename, file_bytes, true);
//                SPList documentLibraryAsList = web.Lists["DocumentLibraryName"];
//                SPListItem itemJustAdded = documentLibraryAsList.GetItemById(newFile.ListItemAllFields.ID);
//                SPContentType documentContentType = documentLibraryAsList.ContentTypes["Document"]; //amend with your document-derived custom Content Type
//                itemJustAdded["ContentTypeId"] = documentContentType.Id;
//                itemJustAdded["Title"] = Title;
//                itemJustAdded.Update();
//                web.AllowUnsafeUpdates = false;
//            }
//        }
//    });
//}


//        public void ReadingFromSystem()
//        {


//            Excel.Application xlapp;
//            Excel.Workbook xlworkbook;
//            Excel.Worksheet xlworksheet;
//            Excel.Range xlrange;

//            string str;
//            int rowcount, colcount;
//            int row = 0;
//            int col = 0;
//            xlapp = new Excel.Application();
//            xlworkbook = xlapp.Workbooks.Open(@"C:\Users\DELL\Desktop\xlsheet.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
//            xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

//            xlrange = xlworksheet.UsedRange;
//            row = xlrange.Rows.Count;
//            col = xlrange.Columns.Count;

//            for (rowcount = 1; rowcount <= row; rowcount++)
//            {
//                for (colcount = 1; colcount <= col; colcount++)
//                {
//                    str = (string)(xlrange.Cells[rowcount, colcount] as Excel.Range).Value2;
//                    Console.WriteLine(str);

//                    xlworkbook.Close(true, null, null);
//                    xlapp.Quit();

//                    Marshal.ReleaseComObject(xlworksheet);
//                    Marshal.ReleaseComObject(xlworkbook);
//                    Marshal.ReleaseComObject(xlapp);

//                }
//                Console.ReadLine();
//            }
//        }

//    }
//}

