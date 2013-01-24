using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    /// <summary>
    /// CellStyle class contains information about cell styling such as cell borders, fill, font, format, text alignment, etc.
    /// </summary>
    public class CellStyle
    {
        internal CellBorder border;
        internal CellFill fill;
        internal CellFont font;
        internal CellAlignment alignment;
        internal String format;


        /// <summary>
        /// Border
        /// </summary>
        public CellBorder Border
        {
            get { return border ?? (border = new CellBorder()); }
            set { border = value; }
        }        

        /// <summary>
        /// Fill
        /// </summary>
        public CellFill Fill
        {
            get { return fill ?? (fill = new CellFill()); }
            set { fill = value; }
        }

        /// <summary>
        /// Font
        /// </summary>
        public CellFont Font
        {
            get { return font ?? (font = new CellFont()); }
            set { font = value; }
        }

        /// <summary>
        /// Alignment
        /// </summary>
        public CellAlignment Alignment
        {
            get { return alignment ?? (alignment = new CellAlignment()); }
            set { alignment = value; }
        }

        /// <summary>
        /// Format
        /// </summary>
        public String Format
        {
            get { return format ?? "General"; }
            set { format = value; }
        }


        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(border, fill, font, alignment, format);
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" /> is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
        public override bool Equals(object obj)
        {
            CellStyle o;
            if (!Util.TryCast(obj, out o))
                return false;
            return border == o.border && fill == o.fill && font == o.font && alignment == o.alignment && format == o.format; 
        }

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(CellStyle c1, CellStyle c2)
        {
            return Util.AreEqual(c1, c2);
        }


        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator !=(CellStyle c1, CellStyle c2)
        {
            return !(c1 == c2);
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

    /// <summary>
    /// CellAlignment class defines cell vertical and horizontal alignment, text indent and rotation, text wrapping and justification.
    /// </summary>
    public class CellAlignment
    {
        public HorizontalAlignment? HAlign { get; set; }
        public VerticalAlignment? VAlign { get; set; }

        /// <summary>
        /// Indent
        /// 1 &eq; 3 spaces
        /// </summary>
        public int Indent { get; set; }

        /// <summary>
        /// Set true to shrink text to fit cell.
        /// </summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>
        /// Text rotation in degrees.
        /// </summary>
        public int TextRotation { get; set; }

        /// <summary>
        /// Set to true to wrap text in multiple lines.
        /// </summary>
        public bool WrapText { get; set; }

        /// <summary>
        /// Set to true to justify last line. Works only if HAlign is set to Justify.
        /// </summary>
        public bool JustifyLastLine { get; set; }


        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(HAlign, VAlign, Indent, ShrinkToFit, TextRotation, WrapText, JustifyLastLine);
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" /> is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
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

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(CellAlignment c1, CellAlignment c2)
        {
            return Util.AreEqual(c1, c2);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator !=(CellAlignment c1, CellAlignment c2)
        {
            return !(c1 == c2);
        }
    }

    public class Color
    {
        public byte a, r, g, b;
        public double tint { get; set; }

        public Color() { }
        
        public Color(byte r, byte g, byte b) : this(255, r, g, b) { }

        public Color(UInt32 argb)
        {
            a = (byte)((argb & 0xFF000000) >> 24);
            r = (byte)((argb & 0x00FF0000) >> 16);
            g = (byte)((argb & 0x0000FF00) >> 8);
            b = (byte)((argb & 0x000000FF) >> 0);
        }

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

        public Color Clone()
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

        public static bool operator ==(Color c1, Color c2)
        {
            return Util.AreEqual(c1, c2);
        }

        public static bool operator !=(Color c1, Color c2)
        {
            return !(c1 == c2);
        }

        public static Color Black { get { return new Color(255, 0, 0, 0); } }
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

        public static bool operator ==(CellFill c1, CellFill c2)
        {
            return Util.AreEqual(c1, c2);
        }

        public static bool operator !=(CellFill c1, CellFill c2)
        {
            return !(c1 == c2);
        }

        public static CellFill Solid(Color color)
        {
            return new CellFill { Foreground = color, Pattern = FillPattern.Solid };
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

        public static bool operator ==(CellBorder c1, CellBorder c2)
        {
            return Util.AreEqual(c1, c2);
        }

        public static bool operator !=(CellBorder c1, CellBorder c2)
        {
            return !(c1 == c2);
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

        public static bool operator ==(BorderEdge c1, BorderEdge c2)
        {
            return Util.AreEqual(c1, c2);
        }

        public static bool operator !=(BorderEdge c1, BorderEdge c2)
        {
            return !(c1 == c2);
        }

        public static BorderEdge Hair()
        {
            return new BorderEdge { Color = Color.Black, Style = BorderStyle.Hair };
        }
        
        public static BorderEdge Hair(Color c)
        {
            return new BorderEdge { Color = c, Style = BorderStyle.Hair };
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

        public static bool operator ==(CellFont c1, CellFont c2)
        {
            return Util.AreEqual(c1, c2);
        }

        public static bool operator !=(CellFont c1, CellFont c2)
        {
            return !(c1 == c2);
        }
    }
}
