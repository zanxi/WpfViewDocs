using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Xml;

namespace ViewAppDocs
{
    class DocxToFlowDocumentConverter : DocReader
    {
        private const string

            // Run properties elements
            ElementBold = "b",
            ElementItalic = "i",
            ElementUnderline = "u",
            ElementStrike = "strike",
            ElementDoubleStrike = "dstrike",
            ElementVerticalAlignment = "vertAlign",
            ElementColor = "color",
            ElementHighlight = "highlight",
            ElementFont = "rFonts",
            ElementFontSize = "sz",
            ElementRightToLeftText = "rtl",

            // Paragraph properties elements
            ElementAlignment = "jc",
            ElementPageBreakBefore = "pageBreakBefore",
            ElementSpacing = "spacing",
            ElementIndentation = "ind",
            ElementShading = "shd",

            // Attributes
            AttributeId = "id",
            AttributeValue = "val",
            AttributeColor = "color",
            AsciiFontFamily = "ascii",
            AttributeSpacingAfter = "after",
            AttributeSpacingBefore = "before",
            AttributeLeftIndentation = "left",
            AttributeRightIndentation = "right",
            AttributeHangingIndentation = "hanging",
            AttributeFirstLineIndentation = "firstLine",
            AttributeFill = "fill";

        private FlowDocument document;
        private TextElement current;
        private bool hasAnyHyperlink;

        public FlowDocument Document
        {
            get { 
                return document; 
            }
        }

        public DocxToFlowDocumentConverter(Stream stream) : base(stream)
        {

        }

        protected override XmlNameTable CreateNameTable()
        {
            var nameTable = base.CreateNameTable();

            nameTable.Add(ElementBold);
            nameTable.Add(ElementItalic);
            nameTable.Add(ElementUnderline);
            nameTable.Add(ElementStrike);
            nameTable.Add(ElementDoubleStrike);
            nameTable.Add(ElementVerticalAlignment);
            nameTable.Add(ElementColor);
            nameTable.Add(ElementHighlight);
            nameTable.Add(ElementFont);
            nameTable.Add(ElementFontSize);
            nameTable.Add(ElementRightToLeftText);
            nameTable.Add(ElementAlignment);
            nameTable.Add(ElementPageBreakBefore);
            nameTable.Add(ElementSpacing);
            nameTable.Add(ElementIndentation);
            nameTable.Add(ElementShading);
            nameTable.Add(AttributeId);
            nameTable.Add(AttributeValue);
            nameTable.Add(AttributeColor);
            nameTable.Add(AsciiFontFamily);
            nameTable.Add(AttributeSpacingAfter);
            nameTable.Add(AttributeSpacingBefore);
            nameTable.Add(AttributeLeftIndentation);
            nameTable.Add(AttributeRightIndentation);
            nameTable.Add(AttributeHangingIndentation);
            nameTable.Add(AttributeFirstLineIndentation);
            nameTable.Add(AttributeFill);

            return nameTable;
        }

        protected override void ReadDocument(XmlReader reader)
        {
            this.document = new FlowDocument();
            this.document.BeginInit();
            this.document.ColumnWidth = double.NaN;

            base.ReadDocument(reader);

            if (this.hasAnyHyperlink) this.document.AddHandler(Hyperlink.RequestNavigateEvent, new RequestNavigateEventHandler((sender, e)=> Process.Start(e.Uri.ToString())));

            this.document.EndInit();
        }

        protected override void ReadParagraph(XmlReader reader)
        {
            using (this.SetCurrent(new Paragraph())) base.ReadParagraph(reader);
            //base.ReadParagraphProperties(reader);
        }

        private IDisposable SetCurrent(TextElement current)
        {
            return new CurrentHandle(this,current);
        }

        protected override void ReadTable(XmlReader reader)
        {
        }

        protected override void ReadParagraphProperties(XmlReader reader)
        {
            while(reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace)
                {
                    var paragraph = (Paragraph)this.current;
                    switch(reader.LocalName)
                    {
                        case ElementAlignment:
                            var textAlignment = ConvertTextAlignment(GetValueAttribute(reader));
                            if (textAlignment.HasValue) paragraph.TextAlignment = textAlignment.Value;
                            break;
                        case ElementPageBreakBefore:
                            paragraph.BreakPageBefore = GetOnOffValueAttribute(reader);                            
                            break;
                        case ElementSpacing:
                            paragraph.Margin = GetSpacing(reader, paragraph.Margin);
                            break;
                        case ElementIndentation:
                            SetParagraphIndent(reader, paragraph);
                            break;
                        case ElementShading:
                            var background  = GetShading(reader);
                            if (background != null) paragraph.Background = background;
                            break;
                    }
                }
            }
        }
        protected override void ReadHyperlink(XmlReader reader)
        {
            var id = reader[AttributeId, RelationshipsNamespace];            
            if(!string.IsNullOrEmpty(id))
            {
                var relationship = this.MainDocumentPart.GetRelationship(id);
                if(relationship.TargetMode == System.IO.Packaging.TargetMode.External)
                {
                    this.hasAnyHyperlink = true;
                    var hyperlink = new Hyperlink() { NavigateUri = relationship.TargetUri };
                    using (this.SetCurrent(hyperlink)) base.ReadHyperlink(reader);
                    return;
                }
            }
            base.ReadHyperlink(reader);
        }

        protected override void ReadRun(XmlReader reader)
        {
            using (SetCurrent(new Span())) base.ReadRun(reader);
        }

        protected override void ReadRunProperties(XmlReader reader)
        {
            while(reader.Read())
            {
                if(reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordprocessingMLNamespace)
                {
                    var inline = (Inline)this.current;
                    switch(reader.LocalName)
                    {
                        case ElementBody:
                            inline.FontWeight = GetOnOffValueAttribute(reader) ? FontWeights.Bold : FontWeights.Normal;
                            break;
                        case ElementItalic:
                            inline.FontStyle = GetOnOffValueAttribute(reader) ? FontStyles.Italic : FontStyles.Normal;
                            break;
                        case ElementUnderline:
                            var underlineTextDecorations = GetUnderlineTextDecorations(reader, inline);
                            if (underlineTextDecorations != null) inline.TextDecorations.Add(underlineTextDecorations);
                            break;
                        case ElementStrike:                            
                            if (GetOnOffValueAttribute(reader)) inline.TextDecorations.Add(TextDecorations.Strikethrough);
                            break;
                        case ElementDoubleStrike:
                            if (GetOnOffValueAttribute(reader))
                            {
                                inline.TextDecorations.Add(new TextDecoration() { Location = TextDecorationLocation.Strikethrough, PenOffset = this.current.FontSize*0.015 });
                                inline.TextDecorations.Add(new TextDecoration() { Location = TextDecorationLocation.Strikethrough, PenOffset = this.current.FontSize * -0.015 });
                            }
                            break;
                        case ElementVerticalAlignment:
                            var baselineAlignment = GetBaselineAlignment(GetValueAttribute(reader));
                            if(baselineAlignment.HasValue)
                            {
                                inline.BaselineAlignment = baselineAlignment.Value;
                                if (baselineAlignment.Value == BaselineAlignment.Subscript || baselineAlignment.Value == BaselineAlignment.Superscript)
                                    inline.FontSize*=0.65;
                            }                            
                            break;
                        case ElementColor:
                            var color = GetColor(GetValueAttribute(reader));
                            if (color.HasValue) inline.Foreground = new SolidColorBrush(color.Value);
                            break;
                        case ElementHighlight:
                            var highlight = GetHighlightColor(GetValueAttribute(reader));
                            if (highlight.HasValue) inline.Background = new SolidColorBrush(highlight.Value);
                            break;
                        case ElementFont:
                            var fontfamily = reader[AsciiFontFamily, WordprocessingMLNamespace];
                            if (!string.IsNullOrEmpty(fontfamily)) inline.FontFamily = (FontFamily)new FontFamilyConverter().ConvertFromString(fontfamily);
                            break;
                        case ElementFontSize:
                            var fontsize = reader[AttributeValue, WordprocessingMLNamespace];
                            if (!string.IsNullOrEmpty(fontsize)) inline.FontSize = uint.Parse(fontsize)*0.66666666666667;
                            break;
                        case ElementRightToLeftText:
                            inline.FlowDirection = (GetOnOffValueAttribute(reader)) ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
                            break;
                    }
                }
            }
        }

        protected override void ReadBreak(XmlReader reader)
        {
            this.AddChild(new LineBreak());
        }

        protected override void ReadTabCharacter(XmlReader reader)
        {
            this.AddChild(new Run("\t"));
        }

        protected override void ReadText(XmlReader reader)
        {
            string txt = reader.ReadString();
            this.AddChild(new Run(txt));
        }





        private TextAlignment? ConvertTextAlignment(string value)
        {
            switch(value)
            {
                case "both":
                    return TextAlignment.Justify;
                case "left":
                    return TextAlignment.Left;
                case "right":
                    return TextAlignment.Right;
                case "Center":
                    return TextAlignment.Center;
                default:
                    return null;
            }

        }

        private string GetValueAttribute(XmlReader reader)
        {
            return reader[AttributeValue, WordprocessingMLNamespace];
        }

        private bool GetOnOffValueAttribute(XmlReader reader)
        {
            var value = GetValueAttribute(reader);
            switch(value)
            {
                case null:
                case "1":
                case "on":
                case "true":
                    return true;
                default:
                    return false;
            }
        }

        private Thickness GetSpacing(XmlReader reader, Thickness margin)
        {
            var after = ConvertTwipsToPixels(reader[AttributeSpacingAfter, WordprocessingMLNamespace]);
            if (after.HasValue) margin.Bottom = after.Value;

            var before = ConvertTwipsToPixels(reader[AttributeSpacingAfter, WordprocessingMLNamespace]);
            if (after.HasValue) margin.Top = before.Value;

            return margin;
        }

        private TextDecorationCollection GetUnderlineTextDecorations(XmlReader reader, Inline inline)
        {
            TextDecoration textDecoration;
            Brush brush;
            var color = GetColor(reader[AttributeColor, WordprocessingMLNamespace]);

            if (color.HasValue) brush = new SolidColorBrush(color.Value);
            else brush = inline.Foreground;

            var textDecorations = new TextDecorationCollection()
            {
                (textDecoration = new TextDecoration
                {
                    Location = TextDecorationLocation.Underline, Pen = new Pen() { Brush = brush }
                }
                )
            };

            switch(GetValueAttribute(reader))
            {
                case "single":
                    break;
                case "double":
                    textDecoration.PenOffset = inline.FontSize * 0.05;
                    textDecoration = textDecoration.Clone();
                    textDecoration.PenOffset = inline.FontSize * -0.05;
                    textDecorations.Add(textDecoration);
                    break;
                case "dotted":
                    textDecoration.Pen.DashStyle = DashStyles.Dot;
                    break;
                case "dash":
                    textDecoration.Pen.DashStyle = DashStyles.Dash;
                    break;
                case "dotDash":
                    textDecoration.Pen.DashStyle = DashStyles.DashDot;
                    break;
                case "dotDotDash":
                    textDecoration.Pen.DashStyle = DashStyles.DashDotDot;
                    break;
                case "none":
                default:
                    return null;
            }
            return textDecorations;
        }

        private double? ConvertTwipsToPixels(string twips)
        {
            if (string.IsNullOrEmpty(twips)) return null;
            else return ConvertTwipsToPixels(double.Parse(twips, CultureInfo.InvariantCulture));
        }

        private double ConvertTwipsToPixels(double twips)
        {           
            return 96d/(72*20)*twips;
        }

        private void SetParagraphIndent(XmlReader reader, Paragraph paragraph)
        {
            var margin = paragraph.Margin;

            var left = ConvertTwipsToPixels(reader[AttributeLeftIndentation, WordprocessingMLNamespace]);
            if (left.HasValue) margin.Left = left.Value;

            var right = ConvertTwipsToPixels(reader[AttributeRightIndentation, WordprocessingMLNamespace]);
            if (right.HasValue) margin.Right = right.Value;

            paragraph.Margin = margin;

            var firstline = ConvertTwipsToPixels(reader[AttributeFirstLineIndentation, WordprocessingMLNamespace]);
            if (firstline.HasValue) paragraph.TextIndent = firstline.Value;

            var hanging = ConvertTwipsToPixels(reader[AttributeHangingIndentation, WordprocessingMLNamespace]);
            if (hanging.HasValue) paragraph.TextIndent -= hanging.Value;
        }

        private Brush GetShading(XmlReader reader)
        {
            var color = GetColor(reader[AttributeFill, WordprocessingMLNamespace]);
            return color.HasValue ? new SolidColorBrush(color.Value) : null;

        }

        private Color? GetColor(string colorString)
        {
            if (string.IsNullOrEmpty(colorString) || colorString == "auto") return null;
            return (Color)ColorConverter.ConvertFromString('#'+colorString);
        }

        private Color? GetHighlightColor(string highlightString)
        {
            if (string.IsNullOrEmpty(highlightString) || highlightString == "auto") return null;
            return (Color)ColorConverter.ConvertFromString(highlightString);
        }

        private BaselineAlignment? GetBaselineAlignment(string verticalAlignmentString)
        {
            switch(verticalAlignmentString)
            {
                case "baseline":
                    return BaselineAlignment.Baseline;
                case "subscript":
                    return BaselineAlignment.Subscript;
                case "superscript":
                    return BaselineAlignment.Superscript;
                default:
                    return null;
            }
        }



        ///////////////////////////////////


        private void AddChild(TextElement textElement)
        {
            ((IAddChild)this.current ?? this.document).AddChild(textElement); 
        }

        private struct CurrentHandle : IDisposable
        {
            private readonly DocxToFlowDocumentConverter converter;
            private readonly TextElement previous;

            public CurrentHandle(DocxToFlowDocumentConverter converter, TextElement current)
            {
                this.converter = converter;
                this.converter.AddChild(current);
                this.previous = this.converter.current;
                this.converter.current = current;
            }

            public void Dispose()
            {
                this.converter.current = this.previous;
            }
        }
                
    }


    
}
