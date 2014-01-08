using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio.Model.Oxml;

namespace Codaxy.Xlio.IO
{
    class Mapper<M, X>
    {
        List<X> outX;
        Dictionary<M, uint> dict;
        Func<M, X> convert;

        public Mapper(Func<M, X> convertor)
        {
            outX = new List<X>();
            dict = new Dictionary<M, uint>();
            this.convert = convertor;
        }

        public uint GetIndex(M m, out X x)
        {
            if (m == null)
            {
                x = default(X);
                return 0;
            }
            uint ind;
            if (!dict.TryGetValue(m, out ind))
            {
                ind = (uint)outX.Count;
                x = convert(m);
                outX.Add(x);
                dict.Add(m, ind);
            }
            x = outX[(int)ind];
            return ind;
        }

        public uint GetIndex(M m)
        {
            X x;
            return GetIndex(m, out x);
        }

        public X[] ToArray() { return outX.ToArray(); }

        public bool Empty { get { return outX.Count == 0; } }

        public void Add(X x)
        {
            outX.Add(x);
        }
    }

    public partial class XlsxFileWriter
    {
        Mapper<CellStyle, CT_Xf> styles;
        Mapper<CellBorder, CT_Border> borders;
        Mapper<CellFill, CT_Fill> fills;
        Mapper<CellFont, CT_Font> fonts;
        Mapper<Color, CT_Color1> colors;
        Mapper<String, CT_NumFmt> numFormats;
        uint numFormatId;

        void ClearStyles()
        {
            styles = new Mapper<CellStyle, CT_Xf>(ConvertStyle);
            styles.Add(new CT_Xf());
            borders = new Mapper<CellBorder, CT_Border>(ConvertBorder);
            borders.Add(new CT_Border());
            fills = new Mapper<CellFill, CT_Fill>(ConvertFill);
            fills.Add(new CT_Fill { Item = new CT_PatternFill { patternType = ST_PatternType.none, patternTypeSpecified = true } });
            fills.Add(new CT_Fill { Item = new CT_PatternFill { patternType = ST_PatternType.gray125, patternTypeSpecified = true } });
            fonts = new Mapper<CellFont, CT_Font>(ConvertFont);
            if (workbook.DefaultFont!=null)
                fonts.Add(ConvertFont(workbook.DefaultFont));
            else
                fonts.Add(new CT_Font { });
            colors = new Mapper<Color, CT_Color1>(ConvertColor);
            numFormats = new Mapper<string, CT_NumFmt>(ConvertNumFormat);
            numFormatId = 164;
        }

        uint RegisterStyle(CellStyle style)
        {
            return styles.GetIndex(style);            
        }

        CT_Xf ConvertStyle(CellStyle style)
        {
            var res =  new CT_Xf();
            if (style.border != null)
            {
                res.borderId = borders.GetIndex(style.border);
                res.borderIdSpecified = true;
                res.applyBorder = true;
                res.applyBorderSpecified = true;
            }
            if (style.fill != null)
            {
                res.fillId = fills.GetIndex(style.fill);
                res.fillIdSpecified = true;
                res.applyFill = true;
                res.applyFillSpecified = true;
            }
            if (style.font != null)
            {
                res.fontId = fonts.GetIndex(style.font);
                res.fontIdSpecified = true;
                res.applyFont = true;
                res.applyFontSpecified = true;
            }
            if (style.alignment != null)
            {

                res.applyAlignmentSpecified = true;
                res.applyAlignment = true;
                res.alignment = ConvertAlignment(style.alignment);
            }
            if (!String.IsNullOrEmpty(style.format))
            {
                res.applyNumberFormatSpecified = true;
                res.applyNumberFormat = true;
                res.numFmtIdSpecified = true;
                uint fmtId;
                if (NumberFormat.TryGetPredefinedFormatId(style.format, out fmtId))
                    res.numFmtId = fmtId;
                else
                {
                    CT_NumFmt fmt;
                    var ind = numFormats.GetIndex(style.format, out fmt);
                    res.numFmtId = fmt.numFmtId;
                }
            }
            return res;
        }

        private CT_CellAlignment ConvertAlignment(CellAlignment a)
        {
            var res = new CT_CellAlignment
            {
                indent = (uint)a.Indent,
                indentSpecified = a.Indent != 0,
                shrinkToFit = a.ShrinkToFit,
                shrinkToFitSpecified = a.ShrinkToFit,
                textRotation = (uint)a.TextRotation,
                textRotationSpecified = a.TextRotation != 0,
                wrapText = a.WrapText,
                wrapTextSpecified = a.WrapText,
                justifyLastLine = a.JustifyLastLine,
                justifyLastLineSpecified = a.JustifyLastLine
            };

            if (a.HAlign.HasValue)
            {
                res.horizontal = (ST_HorizontalAlignment)a.HAlign.Value;
                res.horizontalSpecified = true;
            }

            if (a.VAlign.HasValue)
            {
                res.vertical = (ST_VerticalAlignment)a.VAlign.Value;
                res.verticalSpecified = true;
            }            

            return res;
        }

        CT_Border ConvertBorder(CellBorder border)
        {
            var res = new CT_Border
            {
                left = ConvertBorderPr(border.left),
                right = ConvertBorderPr(border.right),
                top = ConvertBorderPr(border.top),
                bottom = ConvertBorderPr(border.bottom),
                vertical = ConvertBorderPr(border.vertical),
                horizontal = ConvertBorderPr(border.horizontal),
                diagonal = ConvertBorderPr(border.diagonal),
                diagonalDown = border.DiagonalDown.GetValueOrDefault(),
                diagonalDownSpecified = border.DiagonalDown.HasValue,
                diagonalUp = border.DiagonalUp.GetValueOrDefault(),
                diagonalUpSpecified = border.DiagonalUp.HasValue,
                outline = border.Outline
            };
            return res;
        }

        private CT_BorderPr ConvertBorderPr(BorderEdge e)
        {
            if (e == null)
                return new CT_BorderPr();

            return new CT_BorderPr
            {
                style = (ST_BorderStyle)e.Style,
                color = ConvertColor(e.Color)
            };
        }

        private CT_Color1 ConvertColor(Color color)
        {
            if (color == null)
                return null;
            CT_Color1 res = new CT_Color1();
            res.rgb = new[] { color.a, color.r, color.g, color.b };
            return res;
        }

        CT_Font ConvertFont(CellFont font)
        {
            List<object> item = new List<object>();
            List<ItemsChoiceType3> name = new List<ItemsChoiceType3>();

            if (font.Name != null)
            {
                item.Add(new CT_FontName { val = font.Name });
                name.Add(ItemsChoiceType3.name);
            }

            if (font.Size.HasValue)
            {
                item.Add(new CT_FontSize { val = font.Size.Value });
                name.Add(ItemsChoiceType3.sz);
            }

            if (font.Bold)
            {
                item.Add(new CT_BooleanProperty { val = font.Bold });
                name.Add(ItemsChoiceType3.b);
            }

            if (font.Italic)
            {
                item.Add(new CT_BooleanProperty { val = font.Italic });
                name.Add(ItemsChoiceType3.i);
            }

            if (font.Shadow)
            {
                item.Add(new CT_BooleanProperty { val = font.Shadow });
                name.Add(ItemsChoiceType3.shadow);
            }

            if (font.Outline)
            {
                item.Add(new CT_BooleanProperty { val = font.Outline });
                name.Add(ItemsChoiceType3.outline);
            }

            if (font.Strike)
            {
                item.Add(new CT_BooleanProperty { val = font.Strike });
                name.Add(ItemsChoiceType3.strike);
            }

            if (font.Underline.HasValue)
            {
                item.Add(new CT_UnderlineProperty { val = (ST_UnderlineValues)font.Underline.Value });
                name.Add(ItemsChoiceType3.u);
            }

            if (font.Script.HasValue)
            {
                item.Add(new CT_VerticalAlignFontProperty { val = (ST_VerticalAlignRun)font.Script.Value });
                name.Add( ItemsChoiceType3.vertAlign);
            }

            if (font.Color != null)
            {
                item.Add(ConvertColor(font.Color));
                name.Add(ItemsChoiceType3.color);
            }

            return new CT_Font()
            {
                Items = item.ToArray(),
                ItemsElementName = name.ToArray()
            };
        }

        CT_Fill ConvertFill(CellFill fill)
        {
            return new CT_Fill
            {
                Item = new CT_PatternFill
                {
                    patternType = (ST_PatternType)fill.Pattern,
                    patternTypeSpecified = true,
                    fgColor = ConvertColor(fill.Foreground),
                    bgColor = ConvertColor(fill.Background)
                }
            };
        }

        CT_NumFmt ConvertNumFormat(String numFmt)
        {
            var numFmtId = numFormatId;
            numFormatId++;
            return new CT_NumFmt
            {
                numFmtId = numFmtId,
                formatCode = numFmt
            };
        }
    }
}
