using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using FillDocTemplates.Models;
using FillDocTemplates.Word;

namespace FillDocTemplates.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Dossier(int id)
        {
            var doc = WordDocument.CreateFromTemplate(Resource.courrierValeur);

            doc.UpdateMergeField("nomAssure", "Jean Dupont");
            doc.UpdateMergeField("adresse", "1 place de la poste");
            doc.UpdateMergeField("cpville", "31000 TOULOUSE");
            doc.UpdateMergeField("immat", "AF-458-JD");
            doc.UpdateMergeField("valeur", 2500);

            return File(doc.GetDocumentBytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "avis.docx");
        }



        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
