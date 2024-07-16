using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace FileAccess;
public class RecipientModel
{
    public int Sequence { get; set; }
    public string EmployeeCode { get; set; }
    public string EmployeeName { get; set; }
    public  string EmployeeSurname { get; set; }
    public string TaxCertNo { get; set; }
    public string SheetNo { get; set; }
    public string FormSheetNo { get; set; }
    public string Filename { get; set; }
    public string?  EmailAddress { get; set; }
    public string  CompanyCode { get; set; }
    
}
