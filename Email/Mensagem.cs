using OpenPop.Mime;
using System;
using System.Collections.Generic;
using System.Text;

namespace CSharpEmailAnexoINI
{
    public class Mensagem
    {
        public string Id { get; set; }
        public string Assunto { get; set; }
        public string De { get; set; }
        public string Para { get; set; }
        public DateTime Data { get; set; }
        public string ConteudoTexto { get; set; }
        public string ConteudoHtml { get; set; }
        public List<MessagePart> anexos { get; set; }

    }
}
