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

            string username = "sravani.makthala@acuvate.com";
            Console.WriteLine("enter the password");
            SecureString password = GetPassword();


            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ForAssessment/"))
            {


                clientContext.Credentials = new SharePointOnlineCredentials(username, password);
                SP.List spList = clientContext.Web.Lists.GetByTitle("Documents");

                ReadFileName(clientContext);
            }
            Console.ReadLine();
        }
        private static void ReadFileName(ClientContext clientContext)
        {
            string fileName = string.Empty;
            bool isError = true;
            const string fldTitle = "LinkFilename";
            //const string lstDocName = "Documents";
            const string strFolderServerRelativeUrl = "/teams/ForAssessment/Shared%20Documents/";
            string strErrorMsg = string.Empty;
            try
            {
                List list = clientContext.Web.Lists.GetByTitle("Documents");

                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='Recursive'><Query></Query></View>";
                camlQuery.FolderServerRelativeUrl = strFolderServerRelativeUrl;

                SP.ListItemCollection listItems = list.GetItems(camlQuery);

                clientContext.Load(listItems, items => items.Include(i => i[fldTitle]));
                clientContext.ExecuteQuery();
                for (int i = 0; i < listItems.Count; i++)
                {
                    SP.ListItem itemOfInterest = listItems[i];
                    if (itemOfInterest[fldTitle] != null)
                    {
                        fileName = itemOfInterest[fldTitle].ToString();
                        if (i == 0)
                        {
                            ReadExcelData(clientContext, itemOfInterest[fldTitle].ToString());
                        }
                    }
                }
                isError = false;
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
                    //Logging
                }
            }
        }
        private static void ReadExcelData(ClientContext clientContext, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            const string lstDocName = "Documents";
            try
            {
                DataTable dataTable = new DataTable("EmployeeExcelDataTable");
                List list = clientContext.Web.Lists.GetByTitle(lstDocName);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }
                    }
                }
                UpdateSPList(clientContext, dataTable, fileName);
                isError = false;
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
                    //Logging
                }
            }
        }
        private static void UpdateSPList(ClientContext clientContext, DataTable dataTable, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            Int32 count = 0;
            const string lstName = "xlsheet";
            const string lstCol1 = "FilePath";
            const string lstCol2 = "Status";
            const string lstCol3 = "Created By";
            const string lstCol4 = "Department";
            const string lstCol5 = "Uploaded Status";
            const string lstCol6 = "Reason If Failed";
            try
            {
                string fileExtension = ".xlsx";
                string fileNameWithOutExtension = fileName.Substring(0, fileName.Length - fileExtension.Length);
                if (fileNameWithOutExtension.Trim() == lstName)
                {
                    SP.List oList = clientContext.Web.Lists.GetByTitle(fileNameWithOutExtension);
                    foreach (DataRow row in dataTable.Rows)
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem oListItem = oList.AddItem(itemCreateInfo);
                        oListItem[lstCol1] = row[0];
                        oListItem[lstCol2] = row[1];
                        oListItem[lstCol3] = row[2];
                        oListItem[lstCol4] = row[3];
                        oListItem[lstCol5] = row[4];
                        oListItem[lstCol6] = row[5];
                        oListItem.Update();
                        clientContext.ExecuteQuery();
                        count++;
                    }
                }
                else
                {
                    count = 0;
                }
                if (count == 0)
                {
                    Console.Write("Error: List: '" + fileNameWithOutExtension + "' is not found in SharePoint.");
                }
                isError = false;
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
                    // Logging;
                }
            }
        }
        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
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
                    //Logging
                }
            }
            return value;
        }

    }


}


//public void ReadingFromSystem()
//{


//    Excel.Application xlapp;
//    Excel.Workbook xlworkbook;
//    Excel.Worksheet xlworksheet;
//    Excel.Range xlrange;

//    string str;
//    int rowcount, colcount;
//    int row = 0;
//    int col = 0;
//    xlapp = new Excel.Application();
//    xlworkbook = xlapp.Workbooks.Open(@"C:\Users\DELL\Desktop\xlsheet.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
//    xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);

//    xlrange = xlworksheet.UsedRange;
//    row = xlrange.Rows.Count;
//    col = xlrange.Columns.Count;

//    for (rowcount = 1; rowcount <= 1; rowcount++)
//    {
//        for (colcount = 1; colcount <= 1; colcount++)
//        {
//            str = (string)(xlrange.Cells[rowcount, colcount] as Excel.Range).Value2;
//            Console.WriteLine(str);

//            xlworkbook.Close(true, null, null);
//            xlapp.Quit();

//            Marshal.ReleaseComObject(xlworksheet);
//            Marshal.ReleaseComObject(xlworkbook);
//            Marshal.ReleaseComObject(xlapp);

//        }
//        Console.ReadLine();
//    }
//}
