// See https://aka.ms/new-console-template for more information
using FileAccess;
using MailHelperLib;


//SummaryWorksheetAccess workSheetAccess = new();
var file = new FileInfo(@"C:\Users\User\Documents\Jacaranda H\p91\template1.xlsx");
//PDFLocator pDFLocator = new PDFLocator();
//pDFLocator.GetPDFFiles(@"C:\Users\User\Documents\Jacaranda H\p91\Temp_Kenya_P9A_P9B_Tax_De_20230321130711");
//var recipients = await workSheetAccess.ReadDataFromExcelSummaryWorksheet(file);
//foreach(var r in recipients)
//{
//    Console.WriteLine($"{r.Sequence} {r.EmployeeCode} {r.EmployeeName} {r.EmployeeSurname} {r.TaxCertNo}" +
//        $"{r.SheetNo} {r.FormSheetNo} {r.Filename} {r.EmailAddress} {r.CompanyCode}");
//}
EmailHelper helper = new();
await helper.GetMailAddresses(file);
