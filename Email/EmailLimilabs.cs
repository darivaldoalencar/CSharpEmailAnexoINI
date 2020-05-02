using System;
using System.Collections.Generic;
using Limilabs.Client.IMAP;
using Limilabs.Client.POP3;
using Limilabs.Mail;
using Limilabs.Mail.MIME;

namespace CSharpEmailAnexoINI
{
    //Instalar as DLLs: PM> Install-Package Mail.dll
    //https://www.limilabs.com/mail  
    public class EmailLimilabs
    {
        public string hostPop3 { get; set; }
        public string MeuEmaail { get; set; }
        public string MinhaSenha { get; set; }
        public string PathAnexo { get; set; }
        public string hostIMAP { get; set; }

        public void BaixarPop3()
        {
            using (Pop3 pop3 = new Pop3())
            {
                //pop3.Connect("host sem SSL");  
                pop3.ConnectSSL(this.hostPop3);
                pop3.UseBestLogin(this.MeuEmaail, this.MinhaSenha);

                foreach (string uid in pop3.GetAll())
                {
                    IMail email = new MailBuilder()
                        .CreateFromEml(pop3.GetMessageByUID(uid));

                    Console.WriteLine(email.Subject);

                    // salva anexo no disco
                    foreach (MimeData mime in email.Attachments)
                    {
                        mime.Save(string.Concat(this.PathAnexo, mime.SafeFileName));
                    }
                }
                pop3.Close();
            }

        }

        public void BaixarImap()
        {
            using (Imap imap = new Imap())
            {
                //imap.Connect("host sem SSL);   
                imap.ConnectSSL(this.hostIMAP);
                imap.UseBestLogin(this.MeuEmaail, this.MinhaSenha);

                imap.SelectInbox();
                List<long> uids = imap.Search(Flag.All);

                foreach (long uid in uids)
                {
                    IMail email = new MailBuilder()
                        .CreateFromEml(imap.GetMessageByUID(uid));

                    Console.WriteLine(email.Subject);

                    // salva anexo no disco
                    foreach (MimeData mime in email.Attachments)
                    {
                        mime.Save(string.Concat(this.PathAnexo, mime.SafeFileName));
                    }
                }
            }
        }
    }
}
