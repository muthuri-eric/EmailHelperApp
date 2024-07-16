using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAccess;
public class SummaryWorksheetAccess
{
    public async Task<List<RecipientModel>> ReadDataFromExcelSummaryWorksheet(FileInfo file)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        List<RecipientModel> recipients = new();
        if (file != null)
        {
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var ws = package.Workbook.Worksheets["Summary"];

            int row = 6;
            int column = 1;

            while (string.IsNullOrWhiteSpace(ws.Cells[row, column].Value?.ToString()) == false)
            {
                RecipientModel recipientModel = new();
                recipientModel.Sequence = Convert.ToInt32(ws.Cells[row, column].Value);
                recipientModel.EmployeeCode = ws.Cells[row, column + 1].Value.ToString();
                recipientModel.EmployeeName = ws.Cells[row, column + 2].Value.ToString();
                recipientModel.Filename = ws.Cells[row, column + 7].Value.ToString();

                var email = new EmailAddressAttribute();
                string? emailToValidate = ws.Cells[row, column +8].Value?.ToString();
                if(emailToValidate is not null && email.IsValid(emailToValidate) == true)
                {
                    recipientModel.EmailAddress = emailToValidate;
                }
                

                recipientModel.CompanyCode = ws.Cells[row, column + 9].Value.ToString();

                recipients.Add(recipientModel);
                row++;
            }
        }
        return recipients;
    }
}
