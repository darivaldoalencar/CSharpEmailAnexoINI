using CSharpEmailAnexoINI;
using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace Email
{
    class Program
    {

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
                return "C:/ProjetosWeb/CSharpEmailAnexoINI/Email/configuracoes/CONFIGURACOES.INI";
            }
        }

        private string MeuEmaail
        {
            get
            {
                return iniFile.getTAG("EMAIL");
            }
        }

        private bool UseSSL
        {
            get
            {
                return iniFile.getTAG("SSL").ToUpper() == "SIM";
            }
        }

        private int PortaPop3
        {
            get
            {
                return int.Parse(iniFile.getTAG("PORTAPOP3"));
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
                    PathAnexo += "\\";
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
            catch (Exception ex)
            {
                throw new Exception("Erro: " + ex.Message);
            }
        }
        #endregion

        static void Main(string[] args)
        {
            Program p = new Program();
            //EmailLimilabs limabs = new EmailLimilabs()
            //{
            //    hostPop3 = p.hostPop3,
            //    MeuEmaail = p.MeuEmaail,
            //    MinhaSenha = p.MinhaSenha,
            //    PathAnexo = p.PathAnexo,
            //    hostIMAP = p.hostIMAP
            //};
            //limabs.BaixarImap();
            //limabs.BaixarPop3();

            Emails email = new Emails()
            {
                _hostname = p.hostPop3,
                _port = p.PortaPop3,
                _useSsl = p.UseSSL,
                _username = p.MeuEmaail,
                _password = p.MinhaSenha,
                _pathAnexo = p.PathAnexo
            };

            email.DownloadEmail();
        }
    }
}
