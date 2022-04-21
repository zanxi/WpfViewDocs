using System;
using System.Collections.Generic;

//using System.IO.Packaging;
using Packaging = System.IO.Packaging;

using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace ViewAppDocs
{
    public class DocReader : IDisposable
    {
        protected const string
            MainDocumentRelationshipTyoe = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            WordprocessingMLNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",

            ElementDocument = "document",
            ElementBody = "body",

            ElementParagraph = "p",
            ElementTable = "tbl",

            ElementSimpleField = "fldSimple",
            ElementHyperLink = "hyperlink",
            ElementRun = "r",

            ElementBreak = "br",
            ElementTabCharacter = "tab",
            ElementText = "t",

            ElementTableRow = "tr",
            ElementTableCell = "tab",
        //  ElementText = "t",

            ElementParagraphPropeties = "pPr",
            ElementRunPropeties = "rPr";

        protected virtual XmlNameTable CreateNameTable()
        {
            var nameTable = new NameTable();
            nameTable.Add(WordprocessingMLNamespace);
            nameTable.Add(RelationshipsNamespace);
            nameTable.Add(ElementDocument);
            nameTable.Add(ElementBody);
            nameTable.Add(ElementParagraph);
            nameTable.Add(ElementTable);
            nameTable.Add(ElementParagraphPropeties);
            nameTable.Add(ElementSimpleField);
            nameTable.Add(ElementHyperLink);
            nameTable.Add(ElementRun);
            nameTable.Add(ElementBreak);
            nameTable.Add(ElementTabCharacter);
            nameTable.Add(ElementText);
            nameTable.Add(ElementRunPropeties);
            nameTable.Add(ElementTableRow);
            nameTable.Add(ElementTableCell);
            //nameTable.Add();
            //nameTable.Add();
            //nameTable.Add();

            return nameTable;
        }

        private readonly Packaging.Package package;
        private readonly Packaging.PackagePart mainDocumentPart;

        protected Packaging.PackagePart MainDocumentPart
        {
            get { 
                return this.mainDocumentPart; 
            }
        }

        public DocReader(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            this.package = Packaging.Package.Open(stream, FileMode.Open, FileAccess.Read);

            foreach(var rel  in this.package.GetRelationshipsByType(MainDocumentRelationshipTyoe))
            {
                this.mainDocumentPart = package.GetPart(Packaging.PackUriHelper.CreatePartUri(rel.TargetUri));
                break;
            }            
        }

        public void Read()
        {
            using (var mDocStream = this.mainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
            using (var reader = XmlReader.Create(mDocStream, 
                new XmlReaderSettings()
                {
                    NameTable = this.CreateNameTable(),
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true
                }))
            {
                this.ReadMainDocument(reader);
            };
            //  empty fill 2022.04.19
            
        }

        private void ReadXmlSubtree(XmlReader reader, Action<XmlReader> action)
        {
            using (var subtreeReader = reader.ReadSubtree())
            {
                subtreeReader.Read();
                if (action != null) action(subtreeReader);
            }
        }

        private void ReadMainDocument(XmlReader reader)
        {
            while(reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ElementDocument)
                {
                    ReadXmlSubtree(reader, this.ReadDocument);
                    break;
                }
            }
        }

        protected virtual void ReadDocument(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ElementBody)
                {
                    ReadXmlSubtree(reader, this.ReadBody);
                    break;
                }
            }
        }

        private void ReadBody(XmlReader reader)
        {
            while (reader.Read()) this.ReadBlockLevelElement(reader);
        }

        protected virtual void ReadParagraph(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ElementParagraphPropeties)
                    ReadXmlSubtree(reader, this.ReadParagraphProperties);
                else
                    this.ReadInlineLevelElement(reader);
            }
        }

        private void ReadBlockLevelElement(XmlReader reader)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                Action<XmlReader> action = null;

                if (reader.NamespaceURI == WordprocessingMLNamespace)
                    switch (reader.LocalName)
                    {
                        case ElementParagraph:
                            action = this.ReadParagraph;
                            break;

                        case ElementTable:
                            action = this.ReadTable;
                            break;
                    }

                ReadXmlSubtree(reader, action);
            }
        }

        protected virtual void ReadParagraphProperties(XmlReader reader)
        {

        }

        private void ReadSimpleField(XmlReader reader)
        {
            while (reader.Read()) this.ReadInlineLevelElement(reader);
        }

        protected virtual void ReadHyperlink(XmlReader reader)
        {
            while (reader.Read()) this.ReadInlineLevelElement(reader);
        }

        protected virtual void ReadRun(XmlReader reader)
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ElementRunPropeties)
                {
                    ReadXmlSubtree(reader, this.ReadRunProperties);
                }
                else
                {
                    this.ReadRunContentElement(reader);
                }                    
            }
        }

        private void ReadInlineLevelElement(XmlReader reader)
        {
            if(reader.NodeType == XmlNodeType.Element)
            {
                Action<XmlReader> action = null;
                if(reader.NamespaceURI == WordprocessingMLNamespace)
                    switch(reader.LocalName)
                    {
                        case ElementSimpleField:
                            action = this.ReadSimpleField;
                            break;
                        case ElementHyperLink:
                            action = this.ReadHyperlink;
                            break;
                        case ElementRun:
                            action = this.ReadRun;
                            break;
                    };
                ReadXmlSubtree(reader, action);
            }
        }

        protected virtual void ReadRunProperties(XmlReader reader)
        {

        }

        private void ReadRunContentElement(XmlReader reader)
        {
            if(reader.NodeType == XmlNodeType.Element)
            {
                Action<XmlReader> action = null;
                if(reader.NamespaceURI == WordprocessingMLNamespace)
                {
                    switch (reader.LocalName)
                    {
                        case ElementBreak:
                            action = this.ReadBreak;
                            break;
                        case ElementTabCharacter:
                            action = this.ReadTabCharacter;
                            break;
                        case ElementText:
                            action = this.ReadText;
                            break;

                    }
                }
                ReadXmlSubtree(reader, action);
            }
        }

        protected virtual void ReadBreak(XmlReader reader)
        {

        }

        protected virtual void ReadTabCharacter(XmlReader reader)
        {

        }

        protected virtual void ReadText(XmlReader reader)
        {

        }

        protected virtual void ReadTable(XmlReader reader)
        {
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ElementTableRow)
                    ReadXmlSubtree(reader, this.ReadTableRow);
        }

        protected virtual void ReadTableRow(XmlReader reader)
        {
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace && reader.LocalName == ElementTableCell)
                    ReadXmlSubtree(reader, this.ReadTableCell);
        }

        protected virtual void ReadTableCell(XmlReader reader)
        {
            while (reader.Read())
                    ReadBlockLevelElement(reader);
        }

        public void Dispose()
        {
            this.package.Close();
        }


    }


}

