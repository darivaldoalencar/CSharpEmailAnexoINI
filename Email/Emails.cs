using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharpEmailAnexoINI
{
    //Install-Package OpenPop.NET -Version 2.0.6.1120
    public class Emails
    {
        private List<Mensagem> _emails = new List<Mensagem>();
        public string _hostname { get; set; }
        public int _port { get; set; }
        public bool _useSsl { get; set; }
        public string _username { get; set; }
        public string _password { get; set; }
        public string _pathAnexo { get; set; }

        public void DownloadEmail()
        {
            using (var client = new Pop3Client())
            {
                Console.WriteLine("Conectando e-mail: "+ _username);
                client.Connect(_hostname, _port, _useSsl);
                client.Authenticate(_username, _password, AuthenticationMethod.UsernameAndPassword);
                _emails.Clear();

                if (client.Connected)
                {
                    int messageCount = client.GetMessageCount();
                    var Messages = new List<Emails>(messageCount);

                    for (int i = messageCount; i > 0; i--)
                    {
                        var popEmail = client.GetMessage(i);
                        var popText = popEmail.FindFirstPlainTextVersion();
                        var popHtml = popEmail.FindFirstHtmlVersion();

                        string mailText = string.Empty;
                        string mailHtml = string.Empty;
                        if (popText != null)
                            mailText = popText.GetBodyAsText();
                        if (popHtml != null)
                            mailHtml = popHtml.GetBodyAsText();


                        Mensagem mensagem = new Mensagem()
                        {
                            Id = popEmail.Headers.MessageId,
                            Assunto = popEmail.Headers.Subject,
                            De = popEmail.Headers.From.Address,
                            Para = string.Join("; ", popEmail.Headers.To.Select(to => to.Address)),
                            Data = popEmail.Headers.DateSent,
                            ConteudoTexto = mailText,
                            anexos = popEmail.FindAllAttachments(),
                            ConteudoHtml = !string.IsNullOrWhiteSpace(mailHtml) ? mailHtml : mailText
                        };

                        _emails.Add(mensagem);

                        Console.WriteLine("De: "+ mensagem.De);
                        Console.WriteLine("Data: " + mensagem.Data);

                        foreach (var ado in mensagem.anexos)
                        {
                            ado.Save(new FileInfo(Path.Combine(_pathAnexo, ado.FileName)));                            
                            Console.WriteLine("Exportado anexo: " + ado.FileName);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Não consegui conectar");
                }

                Console.ReadKey();
            }
        }
    }
}
