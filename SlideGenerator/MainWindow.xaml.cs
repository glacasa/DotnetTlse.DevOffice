using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;

namespace SlideGenerator
{
    /// <summary>
    /// Cette application prend un "plan" de présentation (txt brut), et crée des slides PowerPoint
    /// </summary>
    public partial class MainWindow : Window
    {
        PowerPoint.Application app = null;
        PowerPoint.Presentation pptx = null;

        public MainWindow()
        {
            InitializeComponent();
        }
        
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ((Button)sender).IsEnabled = false;
                // Ouverture de Powerpoint, et création d'une nouvelle présentation
                app = new PowerPoint.Application();
                pptx = app.Presentations.Add(Office.MsoTriState.msoTrue); // msoFalse permet de travailler en arrière plan sans afficher Office

                // On applique un thème
                string themePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase), "Circuit.thmx");
                pptx.ApplyTheme(themePath);

                // Les 2 layouts de slides qu'on utilise
                var titleLayout = pptx.SlideMaster.CustomLayouts[1];
                var contentLayout = pptx.SlideMaster.CustomLayouts[2];

                // On parcours le texte pour créer les slides
                var content = pptcontent.Text;
                StringReader reader = new StringReader(content);

                int currentIndex = 1;
                string line;
                PowerPoint.Slide slide = null;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    if (line.StartsWith("= "))
                    {
                        // Slide de titre
                        slide = pptx.Slides.AddSlide(currentIndex++, titleLayout);
                        slide.Shapes[1].TextFrame.TextRange.Text = line.Substring(2);
                    }
                    else if (line.StartsWith("# "))
                    {
                        // Le délais n'est là que pour ajouter un effet dramatique pendant la démo, enlevez le pour aller plus vite
                        await Task.Delay(150);
                        // Nouvelle slide de contenu
                        slide = pptx.Slides.AddSlide(currentIndex++, contentLayout);
                        slide.Shapes[1].TextFrame.TextRange.Text = line.Substring(2);
                    }
                    else
                    {
                        // On ajoute la ligne à la dernière slide
                        slide.Shapes[2].TextFrame.TextRange.Text += line + Environment.NewLine;
                    }
                }
         
                // Sauvegarde du document
                // pptx.SaveAs(@"C:\Temp\slides.pptx");
            }
            finally
            {
                // Ici je garde le document ouvert pour la démo, 
                // normalement on pense toujours à fermer Office
                //if (pptx != null) pptx.Close();
                //if (app != null) app.Quit();
                ((Button)sender).IsEnabled = true;
            }
        }
    }
}

