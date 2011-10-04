using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio.Model.Oxml;

namespace Codaxy.Xlio.IO
{
    public partial class XlsxFileReader
    {
        List<CellBorder> borders;
        List<CellStyle> xfs;
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
                       f.Background = GetColor(pf.bgColor);
                       f.Foreground = GetColor(pf.fgColor);
                       f.Pattern = (FillPattern)pf.patternType;
                    }
                    fills.Add(f);
                }

            fonts = new List<CellFont>();
            if (styles.fonts!=null && styles.fonts.font!=null)
                foreach (var font in styles.fonts.font)
                    if (font.ItemsElementName != null && font.Items!=null)
                    {
                        CellFont f = new CellFont();
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
                                        f.Color = GetColor(color);
                                    break;
                            }
                        }
                        fonts.Add(f);
                    }

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
                        entry.Alignment = GetAlignment(xf.alignment);

                    xfs.Add(entry);
                }
        }

        private CellAlignment GetAlignment(CT_CellAlignment al)
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

        private Color GetColor(CT_Color1 c)
        {
            if (c == null)
                return null;

            Color res = new Color();
            if (c.indexedSpecified)
            {
                if (c.indexed <= indexedColors.Count)
                    return indexedColors[(int)c.indexed-1].Clone();
                return null;
            }
            else
            {
                if (c.rgb != null)
                    res = new Color(c.rgb);
            }

            if (res != null)
                res.tint = c.tint;
            return res;                
        }

        private CellBorder GetBorder(int borderId)
        {
            if (borderId < borders.Count)
                return borders[borderId];
            return null;
        }

        private BorderEdge ConvertBorderEdge(CT_BorderPr e)
        {
            if (e == null)
                return null;
            
            BorderEdge edge = new BorderEdge
            {
                Style = (BorderStyle)e.style,
                Color = GetColor(e.color)
            };

            return edge;
        }

        private CellStyle GetStyle(uint index)
        {
            if (index <= 0)
                return null;

            if (index < xfs.Count)
                return xfs[(int)index];

            return null;
        }
    }
}
