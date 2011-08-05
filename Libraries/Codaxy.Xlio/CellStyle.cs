using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class CellStyle
    {
        internal CellBorder border;
        internal CellFill fill;
        internal CellFont font;
        internal CellAlignment alignment;
        internal String format;

        public CellBorder Border
        {
            get { return border ?? (border = new CellBorder()); }
            set { border = value; }
        }
        
        public CellFill Fill
        {
            get { return fill ?? (fill = new CellFill()); }
            set { fill = value; }
        }

        public CellFont Font
        {
            get { return font ?? (font = new CellFont()); }
            set { font = value; }
        }

        public CellAlignment Alignment
        {
            get { return alignment ?? (alignment = new CellAlignment()); }
            set { alignment = value; }
        }

        public String Format
        {
            get { return format ?? "General"; }
            set { format = value; }
        }            
        
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(border, fill, font, alignment, format);
        }

        public override bool Equals(object obj)
        {
            CellStyle o;
            if (!Util.TryCast(obj, out o))
                return false;
            return border == o.border && fill == o.fill && font == o.font && alignment == o.alignment && format == o.format; 
        }        
    }

    public enum HorizontalAlignment
    {
        General,
        Left,
        Center,
        Right,
        Fill,
        Justify,
        CenterContinuous,
        Distributed,
    }

    public enum VerticalAlignment
    {  
        Top,
        Center,
        Bottom,
        Justify,
        Distributed,
    }

    public class CellAlignment
    {
        public HorizontalAlignment? HAlign { get; set; }
        public VerticalAlignment? VAlign { get; set; }

        /// <summary>
        /// 1 <=> 3 spaces
        /// </summary>
        public int Indent { get; set; }
        public bool ShrinkToFit { get; set; }
        public int TextRotation { get; set; }
        public bool WrapText { get; set; }
        public bool JustifyLastLine { get; set; }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(HAlign, VAlign, Indent, ShrinkToFit, TextRotation, WrapText, JustifyLastLine);
        }

        public override bool Equals(object obj)
        {
            CellAlignment o;
            if (!Util.TryCast(obj, out o))
                return false;
            return HAlign == o.HAlign
                && VAlign == o.VAlign
                && Indent == o.Indent
                && TextRotation == o.TextRotation
                && WrapText == o.WrapText
                && JustifyLastLine == o.JustifyLastLine;
        }
    }

    public class Color {

        public byte a, r, g, b;
        public double tint { get; set; }

        public Color() { }
        public Color(params byte[] argb)
        {
            if (argb.Length > 0) 
                a = argb[0];
            if (argb.Length > 1) 
                r = argb[1];
            if (argb.Length > 2) 
                g = argb[2];
            if (argb.Length > 3) 
                b = argb[3];
        }

        internal Color Clone()
        {
            return new Color
            {
                a = a, 
                r = r,
                g = g,
                b = b,
                tint = tint
            };
        }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(a, r, g, b, tint);
        }

        public override bool Equals(object obj)
        {
            Color o;
            if (!Util.TryCast(obj, out o))
                return false;
            return a == o.a && r == o.r && b == o.b && g == o.g && tint == o.tint;
        }        
    }

    public enum FillPattern
    {
        None,
        Solid,
        MediumGray,
        DarkGray,
        LightGray,
        DarkHorizontal,
        DarkVertical,
        DarkDown,
        DarkUp,
        DarkGrid,
        DarkTrellis,
        LightHorizontal,
        LightVertical,
        LightDown,
        LightUp,
        LightGrid,
        LightTrellis,
        Gray125,
        Gray0625
    }

    public class CellFill
    {
        public Color Background { get; set; }
        public Color Foreground { get; set; }
        public FillPattern Pattern { get; set; }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Background, Foreground, Pattern);
        }

        public override bool Equals(object obj)
        {
            CellFill o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Background == o.Background && Foreground == o.Foreground && Pattern == o.Pattern;
        }
    }


    public class CellBorder
    {
        internal BorderEdge left, right, top, bottom, diagonal, horizontal, vertical;

        public BorderEdge Left
        {
            get { return left ?? (left = new BorderEdge()); }
            set { left = value; }
        }
        public BorderEdge Right
        {
            get { return right ?? (right = new BorderEdge()); }
            set { right = value; }
        }
        public BorderEdge Top
        {
            get { return top ?? (top = new BorderEdge()); }
            set { top = value; }
        }
        public BorderEdge Bottom
        {
            get { return bottom ?? (bottom = new BorderEdge()); }
            set { bottom = value; }
        }
        public BorderEdge Diagonal
        {
            get { return diagonal ?? (diagonal = new BorderEdge()); }
            set { diagonal = value; }
        }
        public BorderEdge Horizontal
        {
            get { return horizontal ?? (horizontal = new BorderEdge()); }
            set { horizontal = value; }
        }
        public BorderEdge Vertical
        {
            get { return vertical ?? (vertical = new BorderEdge()); }
            set { vertical = value; }
        }

        public bool? DiagonalUp { get; set; }
        public bool? DiagonalDown { get; set; }
        public bool Outline { get; set; }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(left, right, top, bottom, diagonal, horizontal, vertical, DiagonalUp, DiagonalDown, Outline);
        }

        public override bool Equals(object obj)
        {
            CellBorder o;
            if (!Util.TryCast(obj, out o))
                return false;
            return left == o.left
                && right == o.right
                && top == o.top
                && bottom == o.bottom
                && diagonal == o.diagonal
                && horizontal == o.horizontal
                && vertical == o.vertical
                && DiagonalUp == o.DiagonalUp
                && DiagonalDown == o.DiagonalDown
                && Outline == o.Outline;
        }
    }

    public enum BorderStyle
    {
        None, Thin, Medium, Dashed, Dotted, Thick, Double, Hair, MediumDashed, DashDot, MediumDashDot, DashDotDot, MediumDashDotDot, SlantDashDot,
    }

    public class BorderEdge
    {
        public BorderStyle Style { get; set; }
        public Color Color { get; set; }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Style, Color);
        }

        public override bool Equals(object obj)
        {
            BorderEdge o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Style == o.Style && Color == o.Color;
        }
    }

    public enum FontScript { 
        Baseline,        
        Superscript,
        Subscript,
    }

    public enum FontUnderline
    {
        Single,
        Double,
        SingleAccounting,
        DoubleAccounting,
        None,
    }

    public class CellFont
    {
        public String Name { get; set; }
        public double? Size { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strike { get; set; }
        public FontUnderline? Underline { get; set; }        
        public bool Outline { get; set; }
        public bool Shadow { get; set; }
        public Color Color { get; set; }
        public FontScript? Script { get; set; }

        public CellFont() { Underline = FontUnderline.None; }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Name, Size, Bold, Italic, Strike, Underline, Outline, Shadow, Color, Script);
        }

        public override bool Equals(object obj)
        {
            CellFont o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Name == o.Name
                && Size == o.Size
                && Bold == o.Bold
                && Italic == o.Italic
                && Strike == o.Strike
                && Underline == o.Underline               
                && Outline == o.Outline
                && Shadow == o.Shadow
                && Color == o.Color
                && Script == o.Script;
        }
    }
}
