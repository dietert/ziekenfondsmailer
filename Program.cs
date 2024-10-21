using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using CommandLine;
using MailKit.Net.Smtp;
using MimeKit;
using OfficeOpenXml;
using SelectPdf;
using SendGrid;
using SendGrid.Helpers.Mail;

namespace ZiekenFondsMailer
{
    class Program
    {
        private static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ServicePointManager.ServerCertificateValidationCallback =
                (sender, certificate, chain, sslPolicyErrors) => true;

            Serilog.Information("Ziekenfonds Mailer");
            await Parser.Default.ParseArguments<Options>(args)
                .WithNotParsed(HandleParseError)
                .WithParsedAsync(RunOptions);
        }

        private static async Task RunOptions(Options opts)
        {

            using var package = new ExcelPackage(new FileInfo(opts.ExcelFilePath));
            var pdfBody = File.ReadAllText(opts.HtmlToPdfFilePath);
            var mailBody = File.ReadAllText(opts.MailBodyFilePath);


            foreach (var entry in GetNext(package))
            {
                try
                {
                    var result = await CreateDoc(opts, pdfBody, mailBody, entry, opts.PdfFileName);
                    if (result)
                    {
                        package.Workbook.Worksheets[0].Cells[int.Parse(entry["Rij"]), 4].Value = DateTime.Now.ToOADate();
                        package.Save();
                    }
                }
                catch (Exception exception)
                {
                    Serilog.Error($"Kan geen mail versturen voor rij '{entry["Rij"]}': {exception}");
                }
            }

        }

        static void HandleParseError(IEnumerable<Error> errs)
        {
            foreach (var err in errs)
            {
                Serilog.Error(err.ToString());
            }
        }
        private static IEnumerable<Dictionary<string, string>> GetNext(ExcelPackage package)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var worksheet = package.Workbook.Worksheets[0];
            int row = 2;

            while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].GetValue<string>()))
            {
                if (string.IsNullOrEmpty(worksheet.Cells[row, 2].GetValue<string>()))
                {

                    Serilog.Information($"'{worksheet.Cells[row, 6].GetValue<string>()} {worksheet.Cells[row, 5].GetValue<string>()}' heeft nog niet betaald.");
                    row++;
                    continue;
                }

                if (!string.IsNullOrEmpty(worksheet.Cells[row, 4].GetValue<string>()))
                {

                    Serilog.Information($"Mail naar '{worksheet.Cells[row, 6].GetValue<string>()} {worksheet.Cells[row, 5].GetValue<string>()}' al reeds verstuurd op: '{worksheet.Cells[row, 4].GetValue<DateTime>().ToShortDateString()}'");
                    row++;
                    continue;
                }


                yield return new Dictionary<string, string>
                {
                    {"Rij", row.ToString()},
                    {"Datum", DateTime.Now.ToShortDateString()},
                    {"Seizoen", worksheet.Name.ToUpper().Replace("LEDEN ", "")},
                    {"Voornaam", worksheet.Cells[row, 6].GetValue<string>()},
                    {"Naam", worksheet.Cells[row, 5].GetValue<string>()},
                    {"Straat", worksheet.Cells[row, 8].GetValue<string>()},
                    {"Postcode", worksheet.Cells[row, 9].GetValue<string>()},
                    {"Gemeente", worksheet.Cells[row, 10].GetValue<string>()},
                    {"Geboortedatum", worksheet.Cells[row, 18].GetValue<DateTime>().ToShortDateString()},
                    {"BetaaldBedrag", worksheet.Cells[row, 3].GetValue<string>()},
                    {"DatumBetaling", worksheet.Cells[row, 2].GetValue<DateTime>().ToShortDateString()},
                    {"EmailAdres", worksheet.Cells[row, 12].GetValue<string>()}

                };
                row++;
            }


        }

        private static async Task<bool> CreateDoc(Options opts, string pdfBody, string mailBody,
            Dictionary<string, string> dic, string filename)
        {
            await using var stream = new MemoryStream();

            var doc = new HtmlToPdf().ConvertHtmlString(Replace(pdfBody, dic));

            doc.Save(stream);

            // close pdf document
            doc.Close();

            stream.Seek(0, SeekOrigin.Begin);

            var toEmailAddress = dic["EmailAdres"];
            var toEmailName = dic["Voornaam"] + " " + dic["Naam"];

            return !string.IsNullOrEmpty(opts.SendGridApiKey)
                ? await ExecuteSendGrid(opts, stream, mailBody, dic, filename, toEmailName, toEmailAddress)
                : await ExecuteSmtp(opts, stream, mailBody, dic, filename, toEmailName, toEmailAddress);
        }

        private static string Replace(string input, Dictionary<string, string> values)
        {
            return values.Aggregate(input, (current, kvp) => current.Replace("{{" + kvp.Key + "}}", kvp.Value));
        }

        private static async Task<bool> ExecuteSendGrid(Options opts, MemoryStream stream, string mailBody, Dictionary<string, string> values, string filename, string toEmailName, string toEmailAddress)
        {
            var apiKey = opts.SendGridApiKey;
            var client = new SendGridClient(apiKey);
            var from = new EmailAddress(opts.EmailAddressFrom, opts.EmailNameFrom);
            var subject = opts.MailSubject;
            var to = new EmailAddress(toEmailAddress, toEmailName);
            var plainTextContent = Replace(mailBody, values);

            var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, null);

            msg.AddAttachment(filename, Convert.ToBase64String(stream.ToArray()));

            var response = await client.SendEmailAsync(msg);
            if (response.StatusCode == HttpStatusCode.Accepted)
            {
                Serilog.Information($"Mail verstuurd naar '{to.Name} / {to.Email}'");
                return true;
            }

            var body = await response.DeserializeResponseBodyAsync(response.Body);
            var errormsg = body.FirstOrDefault().ToString();
            Serilog.Error($"Mail naar '{to.Name} / {to.Email}', is niet verstuurd. // HTTP statuscode: '{response.StatusCode}' / {errormsg}");
            return false;
        }


        private static async Task<bool> ExecuteSmtp(Options opts, MemoryStream stream, string mailBody,
            Dictionary<string, string> values, string filename, string toEmailName, string toEmailAddress)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(opts.EmailNameFrom, opts.EmailAddressFrom));
            foreach (var item in EmailAddresses(toEmailAddress))
            {
                message.To.Add(new MailboxAddress(toEmailName, item));
            }
            message.Subject = opts.MailSubject;


            var body = new TextPart("plain")
            {
                Text = Replace(mailBody, values)
            };


            var attachment = new MimePart("application", "pdf")
            {
                Content = new MimeContent(stream),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = filename
            };

            var multipart = new Multipart("mixed") { body, attachment };

            // now set the multipart/mixed as the message body
            message.Body = multipart;

            using (var client = new SmtpClient())
            {
                try
                {
                    var uri = opts.SmtpServer;
                    client.Connect(uri.Host, uri.Port != -1 ? uri.Port : 587, "smtps".Equals(uri.Scheme, StringComparison.InvariantCultureIgnoreCase));

                    if (!string.IsNullOrEmpty(uri.UserInfo))
                    {
                        var username = HttpUtility.UrlDecode(uri.UserInfo.Split(":")[0]);
                        var password = HttpUtility.UrlDecode(uri.UserInfo.Split(":")[1]);
                        client.Authenticate(username, password);
                    }

                    client.Send(message);
                    client.Disconnect(true);
                }
                catch (Exception ex)
                {
                    Serilog.Error($"Mail naar '{toEmailName} / {toEmailAddress}', is niet verstuurd. {ex}");
                    return await Task.FromResult(false);
                }

            }
            Serilog.Information($"Mail verstuurd naar '{toEmailName} / {toEmailAddress}'");
            return await Task.FromResult(true);
        }

        private static IEnumerable<string> EmailAddresses(string emailAddresses)
        {
            if (string.IsNullOrWhiteSpace(emailAddresses)) yield return null;

            var results = emailAddresses.Split(";");

            foreach (var item in results)
            {
                var emailAddress = item.Trim(' ', '\r', '\n');
                if (!string.IsNullOrWhiteSpace(emailAddress)) 
                    yield return emailAddress;
            }
        }
    }

}
