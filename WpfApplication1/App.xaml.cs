using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    /// 

    public class Predio
    {
        public List<String> Entrada = new List<String>();
        public List<String> Saida = new List<String>();

        public List<String> Sorteio(List<String> Entrada)
        {
            List<String> Sorteio = new List<String>();
            int count = 0;
            int selection = 0;

            Random rand = new Random();

            selection = rand.Next(0, 28);
            Sorteio.Add(Entrada[selection]);

            while (count <= 27)
            {
                selection = rand.Next(0, 28);

                if (!Sorteio.Contains(Entrada[selection]))
                {
                    Sorteio.Add(Entrada[selection]);
                }
                count = Sorteio.Count();

            }

            return Sorteio;

        }


    }





    public partial class App : Application
    {
    }
}
