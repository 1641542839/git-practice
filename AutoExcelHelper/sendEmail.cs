using FluentEmail.Core;
using FluentEmail.Smtp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace AutoExcelHelper
{
    class sendEmail
    {
        public async Task SendEmails()
        {
            try
            {
               //config email 
                var emailFrom = ConfigurationManager.AppSettings["EmailFrom"].ToString();
                var emailTO = ConfigurationManager.AppSettings["EmailTo"].ToString();
                var host = ConfigurationManager.AppSettings["Host"].ToString();
                var port = ConfigurationManager.AppSettings["Port"].ToString();
                string[] ef = emailFrom.Split('/');


           

                SmtpClient smtp = new SmtpClient
                {
                    Host = host,
                    Port = Convert.ToInt32(port),
                    EnableSsl = true,
                    UseDefaultCredentials = true,

                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    //set your email credetials
                   // Credentials = new NetworkCredential("xxxx@gmail.com", "xxxx")
                    Credentials = new NetworkCredential(ef[0], ef[1])

                };
                
                String path = Directory.GetCurrentDirectory();
                Email.DefaultSender = new SmtpSender(smtp);

                var attachlist = new List<FluentEmail.Core.Models.Attachment>();
                //var stream = new MemoryStream();
                //var sw = new StreamWriter(stream);
                //sw.Flush();
                //stream.Seek(0, SeekOrigin.Begin);
                path += "\\" + "excelGeneration";
                DirectoryInfo dir = new DirectoryInfo(path);



                foreach (FileInfo flInfo in dir.GetFiles())
                {


                    var attachment = new FluentEmail.Core.Models.Attachment
                    {
                        Data = File.OpenRead(path+"\\"+ flInfo.Name),
                        ContentType = "text/plain",

                        Filename = flInfo.Name

                    };
                    attachlist.Add(attachment);

                }

                //var attachment = new FluentEmail.Core.Models.Attachment
                //{
                //    Data = File.OpenRead(@"C:\Users\allenl\source\repos\AutoExcelHelper\AutoExcelHelper\bin\Debug\excelGeneration\MuliBrand20211108-040628.xls"),

                //    ContentType = "text/plain",

                //    Filename = "MuliBrand20211108-040628.xls"

                //};

                //var attachment1 = new FluentEmail.Core.Models.Attachment
                //{
                //    Data = File.OpenRead(@"C:\Users\allenl\source\repos\AutoExcelHelper\AutoExcelHelper\bin\Debug\excelGeneration\MuliBrand20211108-041427.xls"),

                //    ContentType = "text/plain",

                //    Filename = "MuliBrand20211108-041427.xls"

                //};

               

                 var email = Email
                  //sent your email 
                  .From(ef[0])
                 
                  .To(emailTO)
               
                  .Subject("Sales Report")
                  //.AttachFromFilename($"{Directory.GetCurrentDirectory()}\\excelGeneration\\MuliBrand20211108-034517.xls", null, "MuliBrand20211108-034517.xls")
                  .Attach((IEnumerable<FluentEmail.Core.Models.Attachment>)attachlist)
                  //.Attach(new List<FluentEmail.Core.Models.Attachment> { attachment, attachment1 })
                  //.Attach(IEnumerable<FluentEmail.Core.Models.Attachment> attachlist) 
                  .Body("this is the report contenct, please check in the attachment");


                var result = email.Send();

                if (result.Successful)
                {
                    Console.WriteLine("sending email successful..");
                    Console.WriteLine("complete sending email...");
                }
                else
                {
                    Console.WriteLine("fail to send email");
                }
                

            }
            catch(Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
            }

            
           

        }

    }

}
