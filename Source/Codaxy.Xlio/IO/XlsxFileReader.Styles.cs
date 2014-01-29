using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio.Model.Oxml;
using System.Diagnostics;

namespace Codaxy.Xlio.IO
{
    public partial class XlsxFileReader
    {
        List<CellBorder> borders;
        List<CellStyle> xfs;
        List<CellStyle> dxfs;
        List<CellFill> fills;
        List<Color> indexedColors;
        List<CellFont> fonts;
        Dictionary<uint, String> numFmts;
        
        private void LoadStyles(string path)
        {
            var styles = ReadFile<CT_Stylesheet>(path);

            indexedColors = new List<Color>();

            if (styles.colors != null && styles.colors.indexedColors != null)
                foreach (var color in styles.colors.indexedColors)
                    indexedColors.Add(new Color(color.rgb));

            borders = new List<CellBorder>();
            if (styles.borders!=null && styles.borders.border!=null)
                foreach (var border in styles.borders.border)
                {
                    CellBorder b = new CellBorder
                    {
                        Bottom = ConvertBorderEdge(border.bottom),
                        Right = ConvertBorderEdge(border.right),
                        Left = ConvertBorderEdge(border.left),
                        Top = ConvertBorderEdge(border.top),
                        Diagonal = ConvertBorderEdge(border.diagonal),
                        DiagonalDown = border.diagonalDown,
                        DiagonalUp = border.diagonalUp,
                        Vertical = ConvertBorderEdge(border.vertical),
                        Horizontal = ConvertBorderEdge(border.horizontal),
                        Outline = border.outline
                    };
                    borders.Add(b);
                }     

            fills = new List<CellFill>();
            if (styles.fills != null && styles.fills.fill != null)
                foreach (var fill in styles.fills.fill)
                {
                    CT_PatternFill pf;
                    CellFill f = new CellFill();
                    if (Util.TryCast(fill.Item, out pf))
                    {
                       f.Background = ConvertColor(pf.bgColor);
                       f.Foreground = ConvertColor(pf.fgColor);
                       f.Pattern = (FillPattern)pf.patternType;
                    }
                    fills.Add(f);
                }

            fonts = new List<CellFont>();
            if (styles.fonts!=null && styles.fonts.font!=null)
                foreach (var font in styles.fonts.font)
                {
                    CellFont f = null;
                    if (font.ItemsElementName != null && font.Items != null)
                    {
                        f = new CellFont();
                        for (int i = 0; i < font.ItemsElementName.Length; i++)
                        {
                            var item = font.Items[i];
                            CT_BooleanProperty bp;
                            switch (font.ItemsElementName[i])
                            {
                                case ItemsChoiceType3.name:
                                    CT_FontName name;
                                    if (Util.TryCast(item, out name))
                                        f.Name = name.val;
                                    break;
                                case ItemsChoiceType3.sz:
                                    CT_FontSize size;
                                    if (Util.TryCast(item, out size))
                                        f.Size = size.val;
                                    break;
                                case ItemsChoiceType3.b:
                                    if (Util.TryCast(item, out bp))
                                        f.Bold = bp.val;
                                    break;
                                case ItemsChoiceType3.strike:
                                    if (Util.TryCast(item, out bp))
                                        f.Strike = bp.val;
                                    break;
                                case ItemsChoiceType3.shadow:
                                    if (Util.TryCast(item, out bp))
                                        f.Shadow = bp.val;
                                    break;
                                case ItemsChoiceType3.outline:
                                    if (Util.TryCast(item, out bp))
                                        f.Outline = bp.val;
                                    break;
                                case ItemsChoiceType3.i:
                                    if (Util.TryCast(item, out bp))
                                        f.Italic = bp.val;
                                    break;
                                case ItemsChoiceType3.u:
                                    CT_UnderlineProperty up;
                                    if (Util.TryCast(item, out up))
                                        f.Underline = (FontUnderline)up.val;
                                    break;
                                case ItemsChoiceType3.vertAlign:
                                    CT_VerticalAlignFontProperty vafp;
                                    if (Util.TryCast(item, out vafp))
                                        f.Script = (FontScript)vafp.val;
                                    break;
                                case ItemsChoiceType3.color:
                                    CT_Color1 color;
                                    if (Util.TryCast(item, out color))
                                        f.Color = ConvertColor(color);
                                    break;
                            }
                        }                        
                    }
                    fonts.Add(f);
                }

            if (fonts.Count > 0)
                workbook.DefaultFont = fonts[0];

            numFmts = new Dictionary<uint, string>();
            if (styles.numFmts!=null && styles.numFmts.numFmt!=null)
                foreach (var nfmt in styles.numFmts.numFmt)
                {
                    numFmts[nfmt.numFmtId] = nfmt.formatCode;
                }
            

            xfs = new List<CellStyle>();
            if (styles.cellXfs != null && styles.cellXfs.xf!=null)
                foreach (var xf in styles.cellXfs.xf)
                {
                    var entry = new CellStyle();

                    if (xf.applyBorderSpecified && xf.applyBorder && xf.borderIdSpecified)
                        entry.Border = GetBorder((int)xf.borderId);
                    
                    if (xf.applyFillSpecified && xf.applyFill && xf.fillIdSpecified)
                        entry.Fill = GetFill((int)xf.fillId);

                    if (xf.applyFontSpecified && xf.applyFont && xf.fontIdSpecified)
                        entry.Font = GetFont((int)xf.fontId);
                    
                    if (xf.applyNumberFormatSpecified && xf.applyNumberFormat && xf.numFmtIdSpecified)
                        entry.Format = GetNumFmt(xf.numFmtId);

                    if (xf.applyAlignmentSpecified && xf.applyAlignment && xf.alignment != null)
                        entry.Alignment = ConvertAlignment(xf.alignment);

                    xfs.Add(entry);
                }
           
            dxfs = new List<CellStyle>();
            if (styles.dxfs != null && styles.dxfs.dxf != null)
            {
                foreach (CT_Dxf dxf in styles.dxfs.dxf)
                {
                    var entry = new CellStyle();

                    if (dxf.font != null)
                        entry.Font = ConvertDxfFont(dxf.font);

                    if (dxf.numFmt != null)
                        entry.format = GetNumFmt(dxf.numFmt.numFmtId);

                    if (dxf.fill != null)
                        entry.fill = ConvertDxfCellFill(dxf.fill);

                    if (dxf.alignment != null)
                        entry.alignment = ConvertAlignment(dxf.alignment);

                    if (dxf.border != null)
                        entry.border = ConvertDxfBorder(dxf.border);

                    dxfs.Add(entry);
                }
            }
        }

        private CellFont ConvertDxfFont(CT_Font font)
        {
            var cellFont = new CellFont();
            List<object> items = font.Items.ToList();
            List<ItemsChoiceType3> name = new List<ItemsChoiceType3>();
            int counter = 0;
            foreach (var item in items)
            {
                if (font.ItemsElementName[counter] == ItemsChoiceType3.name)
                {
                    var fontName = item as CT_FontName;
                    if (fontName != null)
                        cellFont.Name = fontName.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.sz)
                {
                    var fontSize = item as CT_FontSize;
                    if (fontSize != null)
                        cellFont.Size = fontSize.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.b)
                {
                    var boldItem = item as CT_BooleanProperty;
                    if (boldItem != null)
                        cellFont.Bold = boldItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.i)
                {
                    var italicItem = item as CT_BooleanProperty;
                    if (italicItem != null)
                        cellFont.Italic = italicItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.shadow)
                {
                    var shadowItem = item as CT_BooleanProperty;
                    if (shadowItem != null)
                        cellFont.Shadow = shadowItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.outline)
                {
                    var outlineItem = item as CT_BooleanProperty;
                    if (outlineItem != null)
                        cellFont.Outline = outlineItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.strike)
                {
                    var strikeItem = item as CT_BooleanProperty;
                    if (strikeItem != null)
                        cellFont.Strike = strikeItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.u)
                {
                    var underlineItem = item as CT_UnderlineProperty;
                    if (underlineItem != null)
                        cellFont.Underline = (FontUnderline)underlineItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.vertAlign)
                {
                    var vertAlignItem = item as CT_VerticalAlignFontProperty;
                    if (vertAlignItem != null)
                        cellFont.Script = (FontScript)vertAlignItem.val;
                }
                if (font.ItemsElementName[counter] == ItemsChoiceType3.color)
                {
                    var colorItem = item as CT_Color1;
                    if (colorItem != null)
                        cellFont.Color = ConvertColor(colorItem);
                }
                counter++;
            }

            return cellFont;
        }

        private CellFill ConvertDxfCellFill(CT_Fill ct_fill)
        {
            CT_PatternFill pf;
            var cellFill = new CellFill();
            if (Util.TryCast(ct_fill.Item, out pf))
            {
                cellFill.Background = ConvertColor(pf.bgColor);
                cellFill.Foreground = ConvertColor(pf.fgColor);
                if (pf.patternType == ST_PatternType.none)
                    pf.patternType = ST_PatternType.solid;
                cellFill.Pattern = (FillPattern)pf.patternType;
            }
            return cellFill;
        }

        private CellBorder ConvertDxfBorder(CT_Border border)
        {
            CellBorder b = new CellBorder
            {
                Bottom = ConvertBorderEdge(border.bottom),
                Right = ConvertBorderEdge(border.right),
                Left = ConvertBorderEdge(border.left),
                Top = ConvertBorderEdge(border.top),
                Diagonal = ConvertBorderEdge(border.diagonal),
                DiagonalDown = border.diagonalDown,
                DiagonalUp = border.diagonalUp,
                Vertical = ConvertBorderEdge(border.vertical),
                Horizontal = ConvertBorderEdge(border.horizontal),
                Outline = border.outline
            };
            
            return b;
        }

        private CellAlignment ConvertAlignment(CT_CellAlignment al)
        {
            CellAlignment a = new CellAlignment
            {
                Indent = al.indentSpecified ? (int)al.indent : 0,
                WrapText = al.wrapTextSpecified ? al.wrapText : false,
                TextRotation = al.textRotationSpecified ? (int)al.textRotation : 0,
                ShrinkToFit = al.shrinkToFitSpecified ? al.shrinkToFit : false,
                VAlign = al.verticalSpecified ? (VerticalAlignment)al.vertical : (VerticalAlignment?)null,
                HAlign = al.horizontalSpecified ? (HorizontalAlignment)al.horizontal : (HorizontalAlignment?)null, 
                JustifyLastLine = al.justifyLastLineSpecified ? al.justifyLastLine : false
            };
            return a;
        }

        private BorderEdge ConvertBorderEdge(CT_BorderPr e)
        {
            if (e == null)
                return null;

            BorderEdge edge = new BorderEdge
            {
                Style = (BorderStyle)e.style,
                Color = ConvertColor(e.color)
            };

            return edge;
        }

        private Color ConvertColor(CT_Color1 c)
        {
            if (c == null)
                return null;

            Color res = new Color();
            if (c.indexedSpecified)
            {
                if (c.indexed < indexedColors.Count)
                    res = indexedColors[(int)c.indexed].Clone();

                res = XlioUtil.GetDefaultIndexedColor(c.indexed);
            }
            else if (c.themeSpecified)
            {
                res = XlioUtil.GetDefaultThemeColor(c.theme);
            }
            else if (c.rgb != null)
            {
                res = new Color(c.rgb);
            }            

            if (res != null)
                res.tint = c.tint;

            return res;                
        }

        private CellStyle GetStyle(uint index)
        {
            //1 based
            if (index < xfs.Count)
                return xfs[(int)index];
            return null;
        }

        private CellStyle GetDxfStyle(uint index)
        {
            if (index < dxfs.Count)
                return dxfs[(int)index];
            return null;
        }
        

        private CellBorder GetBorder(int borderId)
        {
            if (borderId < borders.Count)
                return borders[borderId];
            return null;
        }

        private string GetNumFmt(uint numFmtId)
        {
            string res;
            if (NumberFormat.TryGetPredefinedFormatCode(numFmtId, out res))
                return res;
            if (numFmts.TryGetValue(numFmtId, out res))
                return res;
            return null;
        }

        private CellFont GetFont(int fontId)
        {
            if (fontId < fonts.Count)
                return fonts[fontId];
            return null;
        }

        private CellFill GetFill(int fillId)
        {
            if (fillId < fills.Count)
                return fills[fillId];
            return null;
        }
    }
}
