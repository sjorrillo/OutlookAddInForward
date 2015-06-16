using OutlookAddInForward.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddInForward
{

    public partial class ThisAddIn
    {
        #region Fields

        private const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        #endregion

        #region Events

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void items_ItemAdd(object Item)
        {
            //https://msdn.microsoft.com/en-us/library/bb386179.aspx
            //https://msdn.microsoft.com/en-us/library/cc442767.aspx
            string subject = "Voice Message";
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.MessageClass == "IPM.Note" &&
                 mail.Subject.Equals(subject, StringComparison.OrdinalIgnoreCase) &&
                 mail.Attachments != null && mail.Attachments.Count >= 1)
                {
                    var senderDisplayName = GetSenderDisplayName(mail);
                    var recipientsDisplayName = GetRecipientDisplayName(mail);

                    var llamadaInterna = false;
                    var sourcePhoneNumber = GetSourceNumber(senderDisplayName, out llamadaInterna);
                    var extensionsList = GetExtensions(recipientsDisplayName);

                    extensionsList.ForEach(extension =>
                    {
                        var recipientInfo = GetRecipientInormation(extension);
                        if (!string.IsNullOrEmpty(recipientInfo.Item1))
                        {
                            var voiceMessage = new VoiceMessage
                            {
                                Anexo = extension,
                                UsuarioAnexo = string.IsNullOrEmpty(recipientInfo.Item2) ? " " : recipientInfo.Item2,
                                LlamadaInterna = llamadaInterna,
                                NumeroTelefono = sourcePhoneNumber,
                                FechaLlamada = mail.ReceivedTime
                            };
                            sendMail(mail, recipientInfo.Item1, voiceMessage);
                        }
                    });
                }
            }
        }

        #endregion

        #region Send Email

        private void sendMail(Outlook.MailItem mail, string destinationEmailAddress, VoiceMessage voiceMessage)
        {
            Outlook.MailItem forwardedEmail = Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            forwardedEmail = mail.Forward();

            forwardedEmail.Recipients.Add(destinationEmailAddress);

            forwardedEmail.Subject = "Mensaje de voz recibido";
            var htmlBody = GetHtmlBodyVoiceMessage(voiceMessage);
            forwardedEmail.HTMLBody = htmlBody;
            try
            {
                foreach (Outlook.Account account in Application.Session.Accounts)
                {
                    if (account.AccountType == Outlook.OlAccountType.olPop3 && account.DisplayName.Equals("notificacion.casillavoz@lega.com.pe", StringComparison.OrdinalIgnoreCase))
                    {
                        forwardedEmail.SendUsingAccount = account;
                        break;
                    }
                }

                forwardedEmail.Send();
            }
            catch (Exception ex) {
                System.Diagnostics.Trace.WriteLine(ex.Message, "Outlook Forward Addin");
            }
        }

        #endregion

        #region Get Email Info
        
        private string GetSenderDisplayName(Outlook.MailItem mail)
        {
            if (mail == null)
                throw new ArgumentNullException();

            if (mail.SenderEmailType == "EX")
            {
                Outlook.AddressEntry sender = mail.Sender;
                if (sender != null)
                {
                    //Now we have an AddressEntry representing the Sender
                    if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.Name;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return sender.Name;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return mail.SenderName;
            }
        }


        private List<string> GetRecipientDisplayName(Outlook.MailItem mail)
        {
            if (mail == null)
                throw new ArgumentNullException();

            var recipientsList = new List<string>();
            if (mail.SenderEmailType == "EX")
            {
                Outlook.Recipients recipients = mail.Recipients;
                foreach (Outlook.Recipient recipient in recipients)
                {
                    Outlook.AddressEntry recipientAddress = recipient.AddressEntry;
                    if (recipientAddress == null)
                        continue;

                    //Now we have an AddressEntry representing the Sender
                    if (recipientAddress.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry || recipientAddress.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        Outlook.ExchangeUser exchUser = recipientAddress.GetExchangeUser();
                        if (exchUser != null)
                        {
                            recipientsList.Add(exchUser.Name);
                        }
                        else
                        {
                            continue;
                        }
                    }
                    else
                    {
                        recipientsList.Add(recipientAddress.Name);
                    }
                }
            }
            else
            {
                var recipients = mail.To.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var recipient in recipients)
                {
                    recipientsList.Add(recipient);
                }
            }
            return recipientsList;
        }

        #endregion

        #region Helpers

        private Tuple<string, string> GetRecipientInormation(string extensionNumber)
        {
            var executionPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);
            executionPath = executionPath.Replace("file:\\", "");
            var extensionsFilePath = string.Format("{0}\\extensions.xml", executionPath);

            if (!File.Exists(extensionsFilePath))
                return new Tuple<string, string>(string.Empty, string.Empty);

            var xdocument = XDocument.Load(extensionsFilePath);
            var recipientInfo = (from extension in xdocument.Descendants("extension")
                                 where extension.Attribute("number").Value != null
                                 && extension.Attribute("number").Value.Equals(extensionNumber, StringComparison.OrdinalIgnoreCase)
                                 && extension.Element("owner") != null
                                 select new
                                 {
                                     email = extension.Element("owner").Attribute("email").Value,
                                     name = extension.Element("owner").Attribute("name").Value
                                 }).FirstOrDefault();

            if (recipientInfo == null)
                return new Tuple<string, string>(string.Empty, string.Empty);

            return new Tuple<string, string>(recipientInfo.email, recipientInfo.name);
        }

        private List<string> GetExtensions(List<string> recipientsDisplay)
        {
            var extensionsList = new List<string>();
            recipientsDisplay.ForEach(obj =>
            {
                var recipientSplitted = obj.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                if (recipientSplitted != null)
                {
                    var extension = recipientSplitted[recipientSplitted.GetUpperBound(0)];
                    if (extension.Length > 2)
                    {
                        extension = extension.Substring(1);
                        extension = extension.Substring(0, extension.Length - 1);
                        if (!extensionsList.Any(ext => ext == extension))
                            extensionsList.Add(extension);
                    }
                }

            });
            return extensionsList;
        }

        private string GetSourceNumber(string senderDisplay, out bool llamadaInterna)
        {
            llamadaInterna = false;
            if (senderDisplay.Contains("(") && senderDisplay.Contains(")"))
            {
                llamadaInterna = true;
                var senderSplitted = senderDisplay.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                if (senderSplitted != null)
                {
                    var phoneNumber = senderSplitted[senderSplitted.GetUpperBound(0)];
                    if (phoneNumber.Length > 2)
                    {
                        phoneNumber = phoneNumber.Substring(1);
                        phoneNumber = phoneNumber.Substring(0, phoneNumber.Length - 1);
                        return phoneNumber;
                    }
                }
            }

            return senderDisplay;
        }

        private string GetHtmlBodyVoiceMessage(VoiceMessage voiceMessage)
        {
            var executionPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase);
            executionPath = executionPath.Replace("file:\\", "");
            var templatePath = string.Format("{0}\\Templates\\newVoiceMessage.xsl", executionPath);
            var htmlTemplate = string.Empty;
            var basicTemplate = true;
            if (File.Exists(templatePath))
            {
                var document = GenerateXmlVoiceMessage(voiceMessage);
                var templateResult = EmailTemplate.GenerateEmailBodyFromXslFile(templatePath, document);
                if (templateResult.Success)
                {
                    basicTemplate = false;
                    htmlTemplate = templateResult.CompiledTemplate;
                }
            }

            if (basicTemplate)
            {
                htmlTemplate = "<html>" +
                    "<body style=\"color: red;\">" +
                    "<span style=\"font-variant: small-caps;\">" +
                    "Usted ha recibido un mensaje de voz.</span><br/>" +
                    "<span>" +
                    "En el anexo  : " + voiceMessage.Anexo + "<br/>" +
                    "Del número de teléfono  : " + voiceMessage.NumeroTelefono + "<br/>" +
                    "El : " + voiceMessage.FechaLlamada.ToString("dd/MM/yyyy") + "<br/>" +
                    "Al promediar las :" + voiceMessage.FechaLlamada.ToString("hh:mm tt") +
                    "</span>" +
                    "</body>" +
                    "</html>";
            }
            return htmlTemplate;
        }

        private XDocument GenerateXmlVoiceMessage(VoiceMessage voiceMessage)
        {
            var document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"),
                new XElement("properties",
                    new XElement("anexo", voiceMessage.Anexo),
                    new XElement("usuarioAnexo", voiceMessage.UsuarioAnexo),
                    new XElement("llamadaInterna", voiceMessage.LlamadaInterna ? 1 : 2),
                    new XElement("numeroTelefono", voiceMessage.NumeroTelefono),
                    new XElement("fechaLlamada", voiceMessage.FechaLlamada.ToString("dd/MM/yyyy")),
                    new XElement("horaLlamada", voiceMessage.FechaLlamada.ToString("hh:mm tt"))));
            return document;
        }

        #endregion

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
