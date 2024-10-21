using System;
using CommandLine;

namespace ZiekenFondsMailer
{
    public class Options
    {
        [Option('e', "excelbestand", Required = true, HelpText = "Excel bestandspad.")]
        public string ExcelFilePath { get; set; }

        [Option('h', "htmlbestand", Required = true, HelpText = "Html sjabloon bestandspad.")]
        public string HtmlToPdfFilePath { get; set; }


        [Option('p', "pdfnaam", Required = true, HelpText = "Naam van het pdf bestand als attachment.")]
        public string PdfFileName { get; set; }

        [Option('b', "emailberichtbestand", Required = true, HelpText = "E-mail bericht sjabloon bestandspad.")]
        public string MailBodyFilePath { get; set; }

        [Option('t', "emailtitel", Required = true, HelpText = "E-mail titel sjabloon.")]
        public string MailSubject { get; set; }

        [Option('a', "emailadresvan", Required = true, HelpText = "E-mail adres van.")]
        public string EmailAddressFrom { get; set; }

        [Option('n', "emailnaamvan", Required = true, HelpText = "E-mail naam van.")]
        public string EmailNameFrom { get; set; }

        [Option('s', "smtpserver", Required = false, HelpText = "Smtp server voorbeelden: smtps://dieter.tack%40telenet.be:wachtwoord@smtp.telenet.be:465")]
        public Uri SmtpServer { get; set; }

        [Option('k', "sendgridapikey", Required = false, HelpText = "Send Grid api key.")]
        public string SendGridApiKey { get; set; }

    }
}