using System;
using System.Collections.Generic;
using System.Globalization;
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

        public CellStyle Apply(CellStyle style)
        {
            if (style.alignment != null)
                alignment = style.alignment;

            if (style.border != null)
                border = style.border;

            if (style.fill != null)
                fill = style.fill;

            if (style.font != null)
                font = style.font;

            if (style.format != null)
                format = style.format;

            return this;
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

        public CellStyle Clone()
        {
            return new CellStyle
            {
                border = border,
                format = format,
                font = font,
                fill = fill,
                alignment = alignment
            };

        }
    }

    /// <summary>
    /// Defines horizontal alignment of text
    /// </summary>
    public enum HorizontalAlignment
    {
        /// <summary>
        /// The horizontal alignment is general-aligned. Text data 
        /// is left-aligned. Numbers, dates, and times are right-
        /// aligned. Boolean types are centered. Changing the 
        /// alignment does not change the type of data.  
        /// </summary>
        General,

        /// <summary>
        /// The horizontal alignment is left-aligned, even in Right-to-Left mode. Aligns contents at the left edge of the 
        /// cell. If an indent amount is specified, the contents of 
        /// the cell is indented from the left by the specified 
        /// number of character spaces. The character spaces are 
        /// based on the default font and font size for the 
        /// workbook.
        /// </summary>
        Left,

        /// <summary>
        /// The horizontal alignment is centered, meaning the text is centered across the cell. 
        /// </summary>
        Center,

        /// <summary>
        /// The horizontal alignment is right-aligned, meaning that 
        /// cell contents are aligned at the right edge of the cell, 
        /// even in Right-to-Left mode. 
        /// </summary>
        Right,

        /// <summary>
        /// Indicates that the value of the cell should be filled 
        /// across the entire width of the cell. If blank cells to the 
        /// right also have the fill alignment, they are also filled 
        /// with the value, using a convention similar to 
        /// centerContinuous. 
        /// </summary>
        Fill,

        /// <summary>
        /// The horizontal alignment is justified (flush left and 
        /// right). For each line of text, aligns each line of the 
        /// wrapped text in a cell to the right and left (except the 
        /// last line). If no single line of text wraps in the cell, then 
        /// the text is not justified. 
        /// </summary>
        Justify,

        /// <summary>
        /// The horizontal alignment is centered across multiple 
        /// cells. The information about how many cells to span is 
        /// expressed in the Sheet Part, in the row of the cell in 
        /// question. For each cell that is spanned in the 
        /// alignment, a cell element needs to be written out, with 
        /// the same style Id which references the 
        /// centerContinuous alignment.  
        /// </summary>
        CenterContinuous,

        /// <summary>
        /// I/ndicates that each 'word' in each line of text inside 
        /// the cell is evenly distributed across the width of the 
        /// cell, with flush right and left margins.
        /// </summary>
        Distributed,
    }

    /// <summary>
    /// Defines vertical alignment of text inside a cell.
    /// </summary>
    public enum VerticalAlignment
    {
        /// <summary>
        /// The vertical alignment is aligned-to-top.
        /// </summary>
        Top,

        /// <summary>
        /// The vertical alignment is centered across the height of the cell. 
        /// </summary>
        Center,

        /// <summary>
        /// The vertical alignment is aligned-to-bottom.
        /// </summary>
        Bottom,

        /// <summary>
        /// When text direction is horizontal: the vertical 
        /// alignment of lines of text is distributed vertically, 
        /// where each line of text inside the cell is evenly 
        /// distributed across the height of the cell, with flush top 
        /// and bottom margins.
        /// </summary>
        Justify,

        /// <summary>
        /// When text direction is horizontal: the vertical 
        /// alignment of lines of text is distributed vertically, 
        /// where each line of text inside the cell is evenly 
        /// distributed across the height of the cell, with flush top 
        /// and bottom margins.
        /// </summary>
        Distributed,
    }

    /// <summary>
    /// CellAlignment class defines cell vertical and horizontal alignment, text indent and rotation, text wrapping and justification.
    /// </summary>
    public class CellAlignment
    {
        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        /// <value>
        /// The horizontal alignment.
        /// </value>
        public HorizontalAlignment? HAlign { get; set; }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        /// <value>
        /// The horizontal alignment.
        /// </value>
        public HorizontalAlignment? Horizontal
        {
            get { return HAlign; }
            set { HAlign = value; } 
        }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        /// <value>
        /// The vertical alignment.
        /// </value>
        public VerticalAlignment? VAlign { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        /// <value>
        /// The vertical alignment.
        /// </value>
        public VerticalAlignment? Vertical
        {
            get { return VAlign; }
            set { VAlign = value; }
        }
            

        /// <summary>
        /// Indent
        /// 1 = 3 spaces
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


    /// <summary>
    /// Color class is used to define colors (fill, border, etc.)
    /// </summary>
    public class Color
    {
        /// <summary>
        /// Alpha component (ARGB)
        /// </summary>
        public byte a;

        /// <summary>
        /// Red component (ARGB)
        /// </summary>
        public byte r;

        /// <summary>
        /// Green component (ARGB)
        /// </summary>
        public byte g;

        /// <summary>
        /// Blue component (ARGB)
        /// </summary>
        public byte b;

        /// <summary>
        /// Tint specifies a lighter version of its input color. A 10% tint is 10% of the input color combined with 90% white.
        /// </summary>
        public double tint;

        /// <summary>
        /// Black
        /// </summary>
        public Color() { }

        /// <summary>
        /// RGB color
        /// </summary>
        /// <param name="r"></param>
        /// <param name="g"></param>
        /// <param name="b"></param>
        public Color(byte r, byte g, byte b) : this(255, r, g, b) { }

        /// <summary>
        /// Web color
        /// </summary>
        /// <param name="sharpHex">The web.</param>
        public Color(String sharpHex)
        {
            var l = sharpHex.Length;
            int offset = 0;
            if (sharpHex.StartsWith("#"))
                offset += 1;
            a = 255;
            switch (l - offset)
            {
                case 3:
                    r = byte.Parse(sharpHex[offset + 0].ToString() + sharpHex[offset + 0].ToString(), NumberStyles.HexNumber);
                    g = byte.Parse(sharpHex[offset + 1].ToString() + sharpHex[offset + 1].ToString(), NumberStyles.HexNumber);
                    b = byte.Parse(sharpHex[offset + 2].ToString() + sharpHex[offset + 2].ToString(), NumberStyles.HexNumber);
                    break;
                case 6:
                    r = byte.Parse(sharpHex.Substring(offset + 0, 2), NumberStyles.HexNumber);
                    g = byte.Parse(sharpHex.Substring(offset + 2, 2), NumberStyles.HexNumber);
                    b = byte.Parse(sharpHex.Substring(offset + 4, 2), NumberStyles.HexNumber);
                    break;
                default:
                    throw new InvalidOperationException(String.Format("Invalid color string provided: '{0}'", sharpHex));
            }
        }

        /// <summary>
        /// ARGB
        /// </summary>
        /// <param name="argb"></param>
        public Color(UInt32 argb)
        {
            a = (byte)((argb & 0xFF000000) >> 24);
            r = (byte)((argb & 0x00FF0000) >> 16);
            g = (byte)((argb & 0x0000FF00) >> 8);
            b = (byte)((argb & 0x000000FF) >> 0);
        }

        /// <summary>
        /// ARGB
        /// </summary>
        /// <param name="argb"></param>
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

        /// <summary>
        /// Clone color
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(a, r, g, b, tint);
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
            Color o;
            if (!Util.TryCast(obj, out o))
                return false;
            return a == o.a && r == o.r && b == o.b && g == o.g && tint == o.tint;
        }

        /// <summary>
        /// operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(Color c1, Color c2)
        {
            return Util.AreEqual(c1, c2);
        }

        /// <summary>
        /// operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator !=(Color c1, Color c2)
        {
            return !(c1 == c2);
        }

        /// <summary>
        /// Instance of black color.
        /// </summary>
        public static Color Black { get { return new Color(255, 0, 0, 0); } }
        public static Color White { get { return new Color(255, 255, 255, 255); } }
        public static Color Red { get { return new Color(255, 255, 0, 0); } }
        public static Color Green { get { return new Color(255, 0, 255, 0); } }
        public static Color Blue { get { return new Color(255, 0, 0, 255); } }
        public static Color Yellow { get { return new Color(255, 255, 255, 0); } }
        public static Color Orange { get { return new Color(255, 255, 165, 0); } }
        public static Color Purple { get { return new Color(255, 128, 0, 128); } }
        public static Color Pink { get { return new Color(255, 255, 192, 203); } }
        public static Color Cyan { get { return new Color(255, 0, 255, 255); } }
        public static Color LightBlue { get { return new Color(255, 173, 216, 230); } }
        public static Color Magenta { get { return new Color(255, 255, 0, 255); } }
        public static Color Gray { get { return new Color(255, 190, 190, 190); } }
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

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(CellFill c1, CellFill c2)
        {
            return Util.AreEqual(c1, c2);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
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

        public BorderEdge All
        {
            set
            {
                left = right = top = bottom = value;
            }
        }

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

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(CellBorder c1, CellBorder c2)
        {
            return Util.AreEqual(c1, c2);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator !=(CellBorder c1, CellBorder c2)
        {
            return !(c1 == c2);
        }
    }

    public enum BorderStyle
    {
        /// <summary>
        /// The line style of a border is none (no border visible).
        /// </summary>
        None,

        /// <summary>
        /// The line style of a border is thin.  
        /// </summary>
        Thin,

        /// <summary>
        /// The line style of a border is medium. 
        /// </summary>
        Medium,

        /// <summary>
        /// The line style of a border is dashed.
        /// </summary>
        Dashed,

        /// <summary>
        /// The line style of a border is dotted.
        /// </summary>
        Dotted,

        /// <summary>
        /// The line style of a border is 'thick'.
        /// </summary>
        Thick,

        /// <summary>
        /// The line style of a border is double line.
        /// </summary>
        Double,

        /// <summary>
        /// The line style of a border is hairline. 
        /// </summary>
        Hair,

        /// <summary>
        /// The line style of a border is medium dashed. 
        /// </summary>
        MediumDashed,

        /// <summary>
        /// The line style of a border is dash-dot. 
        /// </summary>
        DashDot,

        /// <summary>
        /// The line style of a border is medium dash-dot.
        /// </summary>
        MediumDashDot,

        /// <summary>
        /// The line style of a border is dash-dot-dot. 
        /// </summary>
        DashDotDot,

        /// <summary>
        /// The line style of a border is medium dash-dot-dot.
        /// </summary>
        MediumDashDotDot,

        /// <summary>
        /// The line style of a border is slant-dash-dot. 
        /// </summary>
        SlantDashDot,
    }

    /// <summary>
    /// BorderEdge contains complete information about border styling
    /// </summary>
    public class BorderEdge
    {
        /// <summary>
        /// Gets or sets the border style.
        /// </summary>
        /// <value>
        /// The style.
        /// </value>
        public BorderStyle Style { get; set; }


        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        /// <value>
        /// The color.
        /// </value>
        public Color Color { get; set; }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Style, Color);
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
            BorderEdge o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Style == o.Style && Color == o.Color;
        }

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(BorderEdge c1, BorderEdge c2)
        {
            return Util.AreEqual(c1, c2);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator !=(BorderEdge c1, BorderEdge c2)
        {
            return !(c1 == c2);
        }

        /// <summary>
        /// Construct black hair border edge
        /// </summary>
        /// <returns></returns>
        public static BorderEdge Hair()
        {
            return new BorderEdge { Color = Color.Black, Style = BorderStyle.Hair };
        }

        /// <summary>
        /// Constructs border edge of given color and hair width.
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static BorderEdge Hair(Color c)
        {
            return new BorderEdge { Color = c, Style = BorderStyle.Hair };
        }
    }

    public enum FontScript
    {
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

    /// <summary>
    /// Defines cell font.
    /// </summary>
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

        /// <summary>
        /// Initializes a new instance of the <see cref="CellFont"/> class.
        /// </summary>
        public CellFont() { Underline = FontUnderline.None; }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Name, Size, Bold, Italic, Strike, Underline, Outline, Shadow, Color, Script);
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

        /// <summary>
        /// operator ==
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator ==(CellFont c1, CellFont c2)
        {
            return Util.AreEqual(c1, c2);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="c2"></param>
        /// <returns></returns>
        public static bool operator !=(CellFont c1, CellFont c2)
        {
            return !(c1 == c2);
        }
    }
}
