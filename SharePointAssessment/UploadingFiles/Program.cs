﻿using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
//using XlReader = Excel;
using Microsoft.Office.SharePoint.Tools;
using System.Net;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;

//using Microsoft.Office.Excel.Server.WebServices;
//using Microsoft.Office.Excel.WebUI;
//using ExcelServiceTest.XLService;



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
                    // Console.Write(ckinfo.KeyChar);
                }
            }
            while (ckinfo.Key != ConsoleKey.Enter);
            return securePassword;
        }
        static void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }
        static void Main(string[] args)
        {

            string username = "sravani.makthala@acuvate.com";
            Console.WriteLine("enter the password");
            SecureString password = GetPassword();

            
            using (var clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/ForAssessment"))
            {


                clientContext.Credentials = new SharePointOnlineCredentials(username, password);
                
                try
                {
                   
                    string filepath ="D:/sravani/sharepoint/SharePointAssessment/Book1.xlsx";

                   
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                    {

                       
                        WorkbookPart wbPart = doc.WorkbookPart;

                        
                        int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count();

                     
                        Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);

                    
                        Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;

                       
                        int wkschildno = 4;


                      
                        SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);


                        
                        Row currentrow = (Row)Rows.ChildElements.GetItem(0);

                       
                        Cell currentcell = (Cell)currentrow.ChildElements.GetItem(0);

                        string currentcellvalue = string.Empty;


                        if (currentcell.DataType != null)
                        {
                            if (currentcell.DataType == CellValues.SharedString)
                            {
                                int id = -1;

                                if (Int32.TryParse(currentcell.InnerText, out id))
                                {
                                    SharedStringItem item = GetSharedStringItemById(wbPart, id);

                                    if (item.Text != null)
                                    {
                                       
                                        currentcellvalue = item.Text.Text;
                                    }
                                    else if (item.InnerText != null)
                                    {
                                        currentcellvalue = item.InnerText;
                                    }
                                    else if (item.InnerXml != null)
                                    {
                                        currentcellvalue = item.InnerXml;
                                    }
                                }
                            }
                        }

                    }
                }
                catch (Exception Ex)
                {

                    Console.WriteLine( Ex.Message);
                }

            }

        }
    }
}
    


