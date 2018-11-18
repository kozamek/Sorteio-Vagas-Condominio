using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Excel;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();

            Predio BlocoA = new Predio();
            Predio BlocoB = new Predio();
            Predio BlocoC = new Predio();
            Predio BlocoD = new Predio();



            BlocoA.Entrada.Add("A11"); BlocoA.Entrada.Add("A12"); BlocoA.Entrada.Add("A13"); BlocoA.Entrada.Add("A14");
            BlocoA.Entrada.Add("A21"); BlocoA.Entrada.Add("A22"); BlocoA.Entrada.Add("A23"); BlocoA.Entrada.Add("A24");
            BlocoA.Entrada.Add("A31"); BlocoA.Entrada.Add("A32"); BlocoA.Entrada.Add("A33"); BlocoA.Entrada.Add("A34");
            BlocoA.Entrada.Add("A41"); BlocoA.Entrada.Add("A42"); BlocoA.Entrada.Add("A43"); BlocoA.Entrada.Add("A44");
            BlocoA.Entrada.Add("A51"); BlocoA.Entrada.Add("A52"); BlocoA.Entrada.Add("A53"); BlocoA.Entrada.Add("A54");
            BlocoA.Entrada.Add("A61"); BlocoA.Entrada.Add("A62"); BlocoA.Entrada.Add("A63"); BlocoA.Entrada.Add("A64");
            BlocoA.Entrada.Add("A71"); BlocoA.Entrada.Add("A72"); BlocoA.Entrada.Add("A73"); BlocoA.Entrada.Add("A74");

            BlocoB.Entrada.Add("B11"); BlocoB.Entrada.Add("B12"); BlocoB.Entrada.Add("B13"); BlocoB.Entrada.Add("B14");
            BlocoB.Entrada.Add("B21"); BlocoB.Entrada.Add("B22"); BlocoB.Entrada.Add("B23"); BlocoB.Entrada.Add("B24");
            BlocoB.Entrada.Add("B31"); BlocoB.Entrada.Add("B32"); BlocoB.Entrada.Add("B33"); BlocoB.Entrada.Add("B34");
            BlocoB.Entrada.Add("B41"); BlocoB.Entrada.Add("B42"); BlocoB.Entrada.Add("B43"); BlocoB.Entrada.Add("B44");
            BlocoB.Entrada.Add("B51"); BlocoB.Entrada.Add("B52"); BlocoB.Entrada.Add("B53"); BlocoB.Entrada.Add("B54");
            BlocoB.Entrada.Add("B61"); BlocoB.Entrada.Add("B62"); BlocoB.Entrada.Add("B63"); BlocoB.Entrada.Add("B64");
            BlocoB.Entrada.Add("B71"); BlocoB.Entrada.Add("B72"); BlocoB.Entrada.Add("B73"); BlocoB.Entrada.Add("B74");

            BlocoC.Entrada.Add("C11"); BlocoC.Entrada.Add("C12"); BlocoC.Entrada.Add("C13"); BlocoC.Entrada.Add("C14");
            BlocoC.Entrada.Add("C21"); BlocoC.Entrada.Add("C22"); BlocoC.Entrada.Add("C23"); BlocoC.Entrada.Add("C24");
            BlocoC.Entrada.Add("C31"); BlocoC.Entrada.Add("C32"); BlocoC.Entrada.Add("C33"); BlocoC.Entrada.Add("C34");
            BlocoC.Entrada.Add("C41"); BlocoC.Entrada.Add("C42"); BlocoC.Entrada.Add("C43"); BlocoC.Entrada.Add("C44");
            BlocoC.Entrada.Add("C51"); BlocoC.Entrada.Add("C52"); BlocoC.Entrada.Add("C53"); BlocoC.Entrada.Add("C54");
            BlocoC.Entrada.Add("C61"); BlocoC.Entrada.Add("C62"); BlocoC.Entrada.Add("C63"); BlocoC.Entrada.Add("C64");
            BlocoC.Entrada.Add("C71"); BlocoC.Entrada.Add("C72"); BlocoC.Entrada.Add("C73"); BlocoC.Entrada.Add("C74");

            BlocoD.Entrada.Add("D11"); BlocoD.Entrada.Add("D12"); BlocoD.Entrada.Add("D13"); BlocoD.Entrada.Add("D14");
            BlocoD.Entrada.Add("D21"); BlocoD.Entrada.Add("D22"); BlocoD.Entrada.Add("D23"); BlocoD.Entrada.Add("D24");
            BlocoD.Entrada.Add("D31"); BlocoD.Entrada.Add("D32"); BlocoD.Entrada.Add("D33"); BlocoD.Entrada.Add("D34");
            BlocoD.Entrada.Add("D41"); BlocoD.Entrada.Add("D42"); BlocoD.Entrada.Add("D43"); BlocoD.Entrada.Add("D44");
            BlocoD.Entrada.Add("D51"); BlocoD.Entrada.Add("D52"); BlocoD.Entrada.Add("D53"); BlocoD.Entrada.Add("D54");
            BlocoD.Entrada.Add("D61"); BlocoD.Entrada.Add("D62"); BlocoD.Entrada.Add("D63"); BlocoD.Entrada.Add("D64");
            BlocoD.Entrada.Add("D71"); BlocoD.Entrada.Add("D72"); BlocoD.Entrada.Add("D73"); BlocoD.Entrada.Add("D74");

            Bloco1Said.ItemsSource = BlocoA.Entrada;
            Bloco2Said.ItemsSource = BlocoB.Entrada;
            Bloco3Said.ItemsSource = BlocoC.Entrada;
            Bloco4Said.ItemsSource = BlocoD.Entrada;

        }

        public void Sorteio_Click(object sender, RoutedEventArgs e)
        {
            Bloco1Said.ItemsSource = null;
            Bloco2Said.ItemsSource = null;
            Bloco3Said.ItemsSource = null;
            Bloco4Said.ItemsSource = null;

            Predio BlocoA = new Predio();
            Predio BlocoB = new Predio();
            Predio BlocoC = new Predio();
            Predio BlocoD = new Predio();

            int i;

            BlocoA.Entrada.Add("A11"); BlocoA.Entrada.Add("A12"); BlocoA.Entrada.Add("A13"); BlocoA.Entrada.Add("A14");
            BlocoA.Entrada.Add("A21"); BlocoA.Entrada.Add("A22"); BlocoA.Entrada.Add("A23"); BlocoA.Entrada.Add("A24");
            BlocoA.Entrada.Add("A31"); BlocoA.Entrada.Add("A32"); BlocoA.Entrada.Add("A33"); BlocoA.Entrada.Add("A34");
            BlocoA.Entrada.Add("A41"); BlocoA.Entrada.Add("A42"); BlocoA.Entrada.Add("A43"); BlocoA.Entrada.Add("A44");
            BlocoA.Entrada.Add("A51"); BlocoA.Entrada.Add("A52"); BlocoA.Entrada.Add("A53"); BlocoA.Entrada.Add("A54");
            BlocoA.Entrada.Add("A61"); BlocoA.Entrada.Add("A62"); BlocoA.Entrada.Add("A63"); BlocoA.Entrada.Add("A64");
            BlocoA.Entrada.Add("A71"); BlocoA.Entrada.Add("A72"); BlocoA.Entrada.Add("A73"); BlocoA.Entrada.Add("A74");

            BlocoB.Entrada.Add("B11"); BlocoB.Entrada.Add("B12"); BlocoB.Entrada.Add("B13"); BlocoB.Entrada.Add("B14");
            BlocoB.Entrada.Add("B21"); BlocoB.Entrada.Add("B22"); BlocoB.Entrada.Add("B23"); BlocoB.Entrada.Add("B24");
            BlocoB.Entrada.Add("B31"); BlocoB.Entrada.Add("B32"); BlocoB.Entrada.Add("B33"); BlocoB.Entrada.Add("B34");
            BlocoB.Entrada.Add("B41"); BlocoB.Entrada.Add("B42"); BlocoB.Entrada.Add("B43"); BlocoB.Entrada.Add("B44");
            BlocoB.Entrada.Add("B51"); BlocoB.Entrada.Add("B52"); BlocoB.Entrada.Add("B53"); BlocoB.Entrada.Add("B54");
            BlocoB.Entrada.Add("B61"); BlocoB.Entrada.Add("B62"); BlocoB.Entrada.Add("B63"); BlocoB.Entrada.Add("B64");
            BlocoB.Entrada.Add("B71"); BlocoB.Entrada.Add("B72"); BlocoB.Entrada.Add("B73"); BlocoB.Entrada.Add("B74");

            BlocoC.Entrada.Add("C11"); BlocoC.Entrada.Add("C12"); BlocoC.Entrada.Add("C13"); BlocoC.Entrada.Add("C14");
            BlocoC.Entrada.Add("C21"); BlocoC.Entrada.Add("C22"); BlocoC.Entrada.Add("C23"); BlocoC.Entrada.Add("C24");
            BlocoC.Entrada.Add("C31"); BlocoC.Entrada.Add("C32"); BlocoC.Entrada.Add("C33"); BlocoC.Entrada.Add("C34");
            BlocoC.Entrada.Add("C41"); BlocoC.Entrada.Add("C42"); BlocoC.Entrada.Add("C43"); BlocoC.Entrada.Add("C44");
            BlocoC.Entrada.Add("C51"); BlocoC.Entrada.Add("C52"); BlocoC.Entrada.Add("C53"); BlocoC.Entrada.Add("C54");
            BlocoC.Entrada.Add("C61"); BlocoC.Entrada.Add("C62"); BlocoC.Entrada.Add("C63"); BlocoC.Entrada.Add("C64");
            BlocoC.Entrada.Add("C71"); BlocoC.Entrada.Add("C72"); BlocoC.Entrada.Add("C73"); BlocoC.Entrada.Add("C74");

            BlocoD.Entrada.Add("D11"); BlocoD.Entrada.Add("D12"); BlocoD.Entrada.Add("D13"); BlocoD.Entrada.Add("D14");
            BlocoD.Entrada.Add("D21"); BlocoD.Entrada.Add("D22"); BlocoD.Entrada.Add("D23"); BlocoD.Entrada.Add("D24");
            BlocoD.Entrada.Add("D31"); BlocoD.Entrada.Add("D32"); BlocoD.Entrada.Add("D33"); BlocoD.Entrada.Add("D34");
            BlocoD.Entrada.Add("D41"); BlocoD.Entrada.Add("D42"); BlocoD.Entrada.Add("D43"); BlocoD.Entrada.Add("D44");
            BlocoD.Entrada.Add("D51"); BlocoD.Entrada.Add("D52"); BlocoD.Entrada.Add("D53"); BlocoD.Entrada.Add("D54");
            BlocoD.Entrada.Add("D61"); BlocoD.Entrada.Add("D62"); BlocoD.Entrada.Add("D63"); BlocoD.Entrada.Add("D64");
            BlocoD.Entrada.Add("D71"); BlocoD.Entrada.Add("D72"); BlocoD.Entrada.Add("D73"); BlocoD.Entrada.Add("D74");

            for (i = 0; i < 10; i++)
                BlocoA.Entrada = BlocoA.Sorteio(BlocoA.Entrada);
            BlocoA.Saida = BlocoA.Sorteio(BlocoA.Entrada);

            for (i = 0; i < 17; i++)
                BlocoB.Entrada = BlocoB.Sorteio(BlocoB.Entrada);
            BlocoB.Saida = BlocoB.Sorteio(BlocoB.Entrada);

            for (i = 0; i < 27; i++)
                BlocoC.Entrada = BlocoC.Sorteio(BlocoC.Entrada);
            BlocoC.Saida = BlocoC.Sorteio(BlocoC.Entrada);

            for (i = 0; i < 33; i++)
                BlocoD.Entrada = BlocoD.Sorteio(BlocoD.Entrada);
            BlocoD.Saida = BlocoD.Sorteio(BlocoD.Entrada);


            Bloco1Said.ItemsSource = BlocoA.Saida;
            Bloco2Said.ItemsSource = BlocoB.Saida;
            Bloco3Said.ItemsSource = BlocoC.Saida;
            Bloco4Said.ItemsSource = BlocoD.Saida;

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            Predio BlocoA = new Predio();
            Predio BlocoB = new Predio();
            Predio BlocoC = new Predio();
            Predio BlocoD = new Predio();

            BlocoA.Entrada.Add("A11"); BlocoA.Entrada.Add("A12"); BlocoA.Entrada.Add("A13"); BlocoA.Entrada.Add("A14");
            BlocoA.Entrada.Add("A21"); BlocoA.Entrada.Add("A22"); BlocoA.Entrada.Add("A23"); BlocoA.Entrada.Add("A24");
            BlocoA.Entrada.Add("A31"); BlocoA.Entrada.Add("A32"); BlocoA.Entrada.Add("A33"); BlocoA.Entrada.Add("A34");
            BlocoA.Entrada.Add("A41"); BlocoA.Entrada.Add("A42"); BlocoA.Entrada.Add("A43"); BlocoA.Entrada.Add("A44");
            BlocoA.Entrada.Add("A51"); BlocoA.Entrada.Add("A52"); BlocoA.Entrada.Add("A53"); BlocoA.Entrada.Add("A54");
            BlocoA.Entrada.Add("A61"); BlocoA.Entrada.Add("A62"); BlocoA.Entrada.Add("A63"); BlocoA.Entrada.Add("A64");
            BlocoA.Entrada.Add("A71"); BlocoA.Entrada.Add("A72"); BlocoA.Entrada.Add("A73"); BlocoA.Entrada.Add("A74");

            BlocoB.Entrada.Add("B11"); BlocoB.Entrada.Add("B12"); BlocoB.Entrada.Add("B13"); BlocoB.Entrada.Add("B14");
            BlocoB.Entrada.Add("B21"); BlocoB.Entrada.Add("B22"); BlocoB.Entrada.Add("B23"); BlocoB.Entrada.Add("B24");
            BlocoB.Entrada.Add("B31"); BlocoB.Entrada.Add("B32"); BlocoB.Entrada.Add("B33"); BlocoB.Entrada.Add("B34");
            BlocoB.Entrada.Add("B41"); BlocoB.Entrada.Add("B42"); BlocoB.Entrada.Add("B43"); BlocoB.Entrada.Add("B44");
            BlocoB.Entrada.Add("B51"); BlocoB.Entrada.Add("B52"); BlocoB.Entrada.Add("B53"); BlocoB.Entrada.Add("B54");
            BlocoB.Entrada.Add("B61"); BlocoB.Entrada.Add("B62"); BlocoB.Entrada.Add("B63"); BlocoB.Entrada.Add("B64");
            BlocoB.Entrada.Add("B71"); BlocoB.Entrada.Add("B72"); BlocoB.Entrada.Add("B73"); BlocoB.Entrada.Add("B74");

            BlocoC.Entrada.Add("C11"); BlocoC.Entrada.Add("C12"); BlocoC.Entrada.Add("C13"); BlocoC.Entrada.Add("C14");
            BlocoC.Entrada.Add("C21"); BlocoC.Entrada.Add("C22"); BlocoC.Entrada.Add("C23"); BlocoC.Entrada.Add("C24");
            BlocoC.Entrada.Add("C31"); BlocoC.Entrada.Add("C32"); BlocoC.Entrada.Add("C33"); BlocoC.Entrada.Add("C34");
            BlocoC.Entrada.Add("C41"); BlocoC.Entrada.Add("C42"); BlocoC.Entrada.Add("C43"); BlocoC.Entrada.Add("C44");
            BlocoC.Entrada.Add("C51"); BlocoC.Entrada.Add("C52"); BlocoC.Entrada.Add("C53"); BlocoC.Entrada.Add("C54");
            BlocoC.Entrada.Add("C61"); BlocoC.Entrada.Add("C62"); BlocoC.Entrada.Add("C63"); BlocoC.Entrada.Add("C64");
            BlocoC.Entrada.Add("C71"); BlocoC.Entrada.Add("C72"); BlocoC.Entrada.Add("C73"); BlocoC.Entrada.Add("C74");

            BlocoD.Entrada.Add("D11"); BlocoD.Entrada.Add("D12"); BlocoD.Entrada.Add("D13"); BlocoD.Entrada.Add("D14");
            BlocoD.Entrada.Add("D21"); BlocoD.Entrada.Add("D22"); BlocoD.Entrada.Add("D23"); BlocoD.Entrada.Add("D24");
            BlocoD.Entrada.Add("D31"); BlocoD.Entrada.Add("D32"); BlocoD.Entrada.Add("D33"); BlocoD.Entrada.Add("D34");
            BlocoD.Entrada.Add("D41"); BlocoD.Entrada.Add("D42"); BlocoD.Entrada.Add("D43"); BlocoD.Entrada.Add("D44");
            BlocoD.Entrada.Add("D51"); BlocoD.Entrada.Add("D52"); BlocoD.Entrada.Add("D53"); BlocoD.Entrada.Add("D54");
            BlocoD.Entrada.Add("D61"); BlocoD.Entrada.Add("D62"); BlocoD.Entrada.Add("D63"); BlocoD.Entrada.Add("D64");
            BlocoD.Entrada.Add("D71"); BlocoD.Entrada.Add("D72"); BlocoD.Entrada.Add("D73"); BlocoD.Entrada.Add("D74");

            Bloco1Said.ItemsSource = BlocoA.Entrada;
            Bloco2Said.ItemsSource = BlocoB.Entrada;
            Bloco3Said.ItemsSource = BlocoC.Entrada;
            Bloco4Said.ItemsSource = BlocoD.Entrada;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {

            int i, ii;

            if (Bloco1Said.ItemsSource != null | Bloco2Said.ItemsSource != null | Bloco3Said.ItemsSource != null | Bloco4Said.ItemsSource != null)
            {
                Microsoft.Office.Interop.Excel.Application XcellApp = new Microsoft.Office.Interop.Excel.Application();
                try
                {

                    XcellApp.Application.Workbooks.Add(Type.Missing);
                    XcellApp.Cells[1, 1] = "BlocoA";

                    ii = 0;

                    for (i = 2; i <= 29; i++)
                    {
                        XcellApp.Cells[i, 1] = Bloco1Said.Items.GetItemAt(ii);
                        ii++;
                    }

                    XcellApp.Cells[1, 3] = "BlocoB";

                    ii = 0;

                    for (i = 2; i <= 29; i++)
                    {
                        XcellApp.Cells[i, 3] = Bloco2Said.Items.GetItemAt(ii);
                        ii++;
                    }

                    XcellApp.Cells[1, 5] = "BlocoC";

                    ii = 0;

                    for (i = 2; i <= 29; i++)
                    {
                        XcellApp.Cells[i, 5] = Bloco3Said.Items.GetItemAt(ii);
                        ii++;
                    }

                    XcellApp.Cells[1, 7] = "BlocoD";

                    ii = 0;

                    for (i = 2; i <= 29; i++)
                    {
                        XcellApp.Cells[i, 7] = Bloco4Said.Items.GetItemAt(ii);
                        ii++;
                    }

                    XcellApp.Visible = true;

                }

                catch (Exception ex)
                {
                    MessageBox.Show("Erro :" + ex.Message);
                    XcellApp.Quit();

                }

            }
            else { MessageBoxResult result = MessageBox.Show("Favor Realizar o Sorteio Primeiro"); }


        }


    }

}

