using FileAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace MailHelperLib;
public class EmailHelper
{
    
    public async Task GetMailAddresses(FileInfo file)
    {
        //MailAddress mailAddress = new()
        SummaryWorksheetAccess summaryWorksheetAccess = new();
        var recipients = await summaryWorksheetAccess.ReadDataFromExcelSummaryWorksheet(file);
        foreach(var r in recipients)
        {
            if(r.EmailAddress is not null)
            {
                MailAddress to = new(r.EmailAddress, r.EmployeeName);
                MailAddress from = new("eric@maarifa.co.ke");//pass the email from here

                MailMessage message = new MailMessage(from, to);

                message.Subject = "P9 for Tax Year 2022";
                message.Body = $"Good Day {r.EmployeeName} Please find attached your P9 form." +
                    $" Use your ID Number to open the file";

                Attachment fileAttachment = new Attachment($"C:\\test\\{r.Filename}");
                message.Attachments.Add(fileAttachment);

                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = "smtp.gmail.com";
                smtpClient.Port = 587;
                smtpClient.EnableSsl = true;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential("eric@maarifa.co.ke", "6555eric");
                smtpClient.Send(message);
            }
            
        }
    }
}
