using Limilabs.Client.IMAP;
using Limilabs.Client.POP3;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using System;
using System.Collections.Generic;
using System.IO;

namespace Email
{
    class Program
    {
        //Instalar as DLLs: PM> Install-Package Mail.dll
        //https://www.limilabs.com/mail        

        #region propriedades
        private IniFile iniFile
        {
            get
            {                
                return new IniFile(pathINI);
            }
        }

        private static string pathINI
        {
            get
            {                
                return "C:\\ProjetosWeb\\CSharpEmailAnexoINI\\Email\\configuracoes\\CONFIGURACOES.INI";
            }
        }

        private string MeuEmaail
        {
            get
            {
                return iniFile.getTAG("EMAIL");
            }
        }

        private string MinhaSenha
        {
            get
            {
                return iniFile.getTAG("SENHA");
            }
        }
        private string PathAnexo
        {
            get
            {
                string PathAnexo = iniFile.getTAG("PATHANEXO");
                if (!PathAnexo[PathAnexo.Length - 1].ToString().Equals("\\"))
                {
                    PathAnexo +=  "\\";
                }
                CriarPastasAmbiente(PathAnexo);
                return PathAnexo;
            }
        }
        private string hostPop3
        {
            get
            {
                return iniFile.getTAG("HOSTPOP3");
            }
        }

        private string hostIMAP
        {
            get
            {
                return iniFile.getTAG("HOSTIMAP");
            }
        }
        #endregion

        #region métodos
        private void BaixarPop3()
        {
            using (Pop3 pop3 = new Pop3())
            {
                //pop3.Connect("host sem SSL");  
                pop3.ConnectSSL(hostPop3);
                pop3.UseBestLogin(MeuEmaail, MinhaSenha);

                foreach (string uid in pop3.GetAll())
                {
                    IMail email = new MailBuilder()
                        .CreateFromEml(pop3.GetMessageByUID(uid));

                    Console.WriteLine(email.Subject);

                    // salva anexo no disco
                    foreach (MimeData mime in email.Attachments)
                    {
                        mime.Save(string.Concat(PathAnexo, mime.SafeFileName));
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
                imap.ConnectSSL(hostIMAP);
                imap.UseBestLogin(MeuEmaail, MinhaSenha);

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
                        mime.Save(string.Concat(PathAnexo, mime.SafeFileName));
                    }
                }
            }
        }        

        private bool CriarPastasAmbiente(string pasta)
        {
            try
            {
                if (!Directory.Exists(pasta))
                {
                    Directory.CreateDirectory(pasta);
                }
                return true;
            }
            catch(Exception ex)
            {
                throw new Exception("Erro: " + ex.Message);
            }
        }
        #endregion

        static void Main(string[] args)
        {            
            var p = new Program();
            
            //a versão pop3 funciona também
            //p.BaixarPop3();

            p.BaixarImap();            
        }
    }
}
