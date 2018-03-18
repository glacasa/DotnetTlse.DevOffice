using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FillDocTemplates.Word
{
    public class WordDocument
    {
        public WordprocessingDocument Document { get; private set; }
        private MemoryStream DocumentStream { get; set; }

        private WordDocument() { }

        /// <summary>
        /// Ouvre un document ou template Word, et crée un nouveau modèle basé dessus
        /// </summary>
        public static WordDocument CreateFromTemplate(byte[] byteArray)
        {
            // Ici je travaille avec le document en mémoire, et pas sur un document enregistré
            // Je garde donc toujours une référence vers le MemoryStream
            var stream = new MemoryStream(byteArray.Length);
            stream.Write(byteArray, 0, byteArray.Length);

            var document = WordprocessingDocument.Open(stream, true);
            // Si on ouvre un template (.dotx), on doit le convertir en document normal (.docx)
            document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            document.Save();

            return new WordDocument()
            {
                Document = document,
                DocumentStream = stream
            };
        }



        private Regex MergeFieldRegex = new Regex(@"MERGEFIELD\s+([a-zA-Z0-9_-]+)", RegexOptions.Compiled);


        public Dictionary<string, List<Text>> GetMergeFields()
        {
            // En OpenXML, un champ de fusion est un objet SimpleField, qui contient l'instruction "MERGEFIELD [nom du champ de fusion]"

            var fieldsCollection = new Dictionary<string, List<Text>>();
            MainDocumentPart mainPart = Document.MainDocumentPart;

            // On recherche les objets Simple fields
            var simpleFields = mainPart.RootElement.Descendants<SimpleField>();
            foreach (var field in simpleFields)
            {
                // Le SimpleField est-il un champ de fusion ? (MERGEFIELD)
                var match = MergeFieldRegex.Match(field.Instruction.InnerText);
                if (match.Success)
                {
                    // Ok, on peut sauvegarder le champ de fusion, à partir de son nom
                    var fieldName = match.Groups[1].Value;

                    var valueContainer = field.Descendants<Text>().SingleOrDefault();
                    if (valueContainer != null)
                    {
                        if (!fieldsCollection.ContainsKey(fieldName))
                        {
                            fieldsCollection.Add(fieldName, new List<Text>());
                        }
                        fieldsCollection[fieldName].Add(valueContainer);
                    }
                }
            }

            return fieldsCollection;
        }

        public void UpdateMergeField(string fieldName, object value)
        {
            var field = GetMergeFields()[fieldName];
            // field est la liste des objets `Text` du document associés aux champs de fusion

            foreach (var textNode in field)
            {
                textNode.Text = value.ToString();
            }

            Document.Save();
        }

        public byte[] GetDocumentBytes()
        {
            // Il faut fermer le document pour avoir le MemoryStream à jour
            Document.Close();
            return DocumentStream.ToArray();
        }


    }
}
