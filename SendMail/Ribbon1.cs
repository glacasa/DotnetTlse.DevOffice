using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace AddSlide
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // On trouve le powerpoint ouvert, et on ajoute une slide à la fin
            var app = Globals.ThisAddIn.Application;
            var pptx = app.ActivePresentation;
            if (pptx != null)
            {
                var titleLayout = pptx.SlideMaster.CustomLayouts[1];
                var qSlide = pptx.Slides.AddSlide(pptx.Slides.Count + 1, titleLayout);

                qSlide.Shapes[1].TextFrame.TextRange.Text = "Questions ?";
            }
        }
    }
}
