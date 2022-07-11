using System;
using System.Collections.Generic;
using System.Text;

namespace AlertaResina
{
    public class EmailDto
    {
        public string[] Destinatarios { get; set; }
        public string Asunto { get; set; }
        public string Mensaje { get; set; }

    }
}
