using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Model.Oxml
{
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("styleSheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_Stylesheet
    {

        private CT_NumFmts numFmtsField;

        private CT_Fonts fontsField;

        private CT_Fills fillsField;

        private CT_Borders bordersField;

        private CT_CellStyleXfs cellStyleXfsField;

        private CT_CellXfs cellXfsField;

        private CT_CellStyles cellStylesField;

        private CT_Dxfs dxfsField;

        private CT_TableStyles tableStylesField;

        private CT_Colors colorsField;

        private CT_ExtensionList extLstField;

        /// <remarks/>
        public CT_NumFmts numFmts
        {
            get
            {
                return this.numFmtsField;
            }
            set
            {
                this.numFmtsField = value;
            }
        }

        /// <remarks/>
        public CT_Fonts fonts
        {
            get
            {
                return this.fontsField;
            }
            set
            {
                this.fontsField = value;
            }
        }

        /// <remarks/>
        public CT_Fills fills
        {
            get
            {
                return this.fillsField;
            }
            set
            {
                this.fillsField = value;
            }
        }

        /// <remarks/>
        public CT_Borders borders
        {
            get
            {
                return this.bordersField;
            }
            set
            {
                this.bordersField = value;
            }
        }

        /// <remarks/>
        public CT_CellStyleXfs cellStyleXfs
        {
            get
            {
                return this.cellStyleXfsField;
            }
            set
            {
                this.cellStyleXfsField = value;
            }
        }

        /// <remarks/>
        public CT_CellXfs cellXfs
        {
            get
            {
                return this.cellXfsField;
            }
            set
            {
                this.cellXfsField = value;
            }
        }

        /// <remarks/>
        public CT_CellStyles cellStyles
        {
            get
            {
                return this.cellStylesField;
            }
            set
            {
                this.cellStylesField = value;
            }
        }

        /// <remarks/>
        public CT_Dxfs dxfs
        {
            get
            {
                return this.dxfsField;
            }
            set
            {
                this.dxfsField = value;
            }
        }

        /// <remarks/>
        public CT_TableStyles tableStyles
        {
            get
            {
                return this.tableStylesField;
            }
            set
            {
                this.tableStylesField = value;
            }
        }

        /// <remarks/>
        public CT_Colors colors
        {
            get
            {
                return this.colorsField;
            }
            set
            {
                this.colorsField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_NumFmts
    {

        private CT_NumFmt[] numFmtField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("numFmt")]
        public CT_NumFmt[] numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_NumFmt
    {

        private uint numFmtIdField;

        private string formatCodeField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string formatCode
        {
            get
            {
                return this.formatCodeField;
            }
            set
            {
                this.formatCodeField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Fonts
    {

        private CT_Font[] fontField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("font")]
        public CT_Font[] font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Font
    {

        private object[] itemsField;

        private ItemsChoiceType3[] itemsElementNameField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("b", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("charset", typeof(CT_IntProperty))]
        [System.Xml.Serialization.XmlElementAttribute("color", typeof(CT_Color1))]
        [System.Xml.Serialization.XmlElementAttribute("condense", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("extend", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("family", typeof(CT_IntProperty))]
        [System.Xml.Serialization.XmlElementAttribute("i", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("name", typeof(CT_FontName))]
        [System.Xml.Serialization.XmlElementAttribute("outline", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("scheme", typeof(CT_FontScheme1))]
        [System.Xml.Serialization.XmlElementAttribute("shadow", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("strike", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("sz", typeof(CT_FontSize))]
        [System.Xml.Serialization.XmlElementAttribute("u", typeof(CT_UnderlineProperty))]
        [System.Xml.Serialization.XmlElementAttribute("vertAlign", typeof(CT_VerticalAlignFontProperty))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public ItemsChoiceType3[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_BooleanProperty
    {

        private bool valField;

        public CT_BooleanProperty()
        {
            this.valField = true;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_IntProperty
    {

        private int valField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName = "CT_Color", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Color1
    {

        private bool autoField;

        private bool autoFieldSpecified;

        private uint indexedField;

        private bool indexedFieldSpecified;

        private byte[] rgbField;

        private uint themeField;

        private bool themeFieldSpecified;

        private double tintField;

        public CT_Color1()
        {
            this.tintField = 0;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool auto
        {
            get
            {
                return this.autoField;
            }
            set
            {
                this.autoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool autoSpecified
        {
            get
            {
                return this.autoFieldSpecified;
            }
            set
            {
                this.autoFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint indexed
        {
            get
            {
                return this.indexedField;
            }
            set
            {
                this.indexedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool indexedSpecified
        {
            get
            {
                return this.indexedFieldSpecified;
            }
            set
            {
                this.indexedFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType = "hexBinary")]
        public byte[] rgb
        {
            get
            {
                return this.rgbField;
            }
            set
            {
                this.rgbField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint theme
        {
            get
            {
                return this.themeField;
            }
            set
            {
                this.themeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool themeSpecified
        {
            get
            {
                return this.themeFieldSpecified;
            }
            set
            {
                this.themeFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double tint
        {
            get
            {
                return this.tintField;
            }
            set
            {
                this.tintField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_FontName
    {

        private string valField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName = "CT_FontScheme", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_FontScheme1
    {

        private ST_FontScheme valField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_FontScheme val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_FontScheme
    {

        /// <remarks/>
        none,

        /// <remarks/>
        major,

        /// <remarks/>
        minor,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_FontSize
    {

        private double valField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_UnderlineProperty
    {

        private ST_UnderlineValues valField;

        public CT_UnderlineProperty()
        {
            this.valField = ST_UnderlineValues.single;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_UnderlineValues.single)]
        public ST_UnderlineValues val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_UnderlineValues
    {

        /// <remarks/>
        single,

        /// <remarks/>
        @double,

        /// <remarks/>
        singleAccounting,

        /// <remarks/>
        doubleAccounting,

        /// <remarks/>
        none,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_VerticalAlignFontProperty
    {

        private ST_VerticalAlignRun valField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_VerticalAlignRun val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_VerticalAlignRun
    {

        /// <remarks/>
        baseline,

        /// <remarks/>
        superscript,

        /// <remarks/>
        subscript,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType3
    {

        /// <remarks/>
        b,

        /// <remarks/>
        charset,

        /// <remarks/>
        color,

        /// <remarks/>
        condense,

        /// <remarks/>
        extend,

        /// <remarks/>
        family,

        /// <remarks/>
        i,

        /// <remarks/>
        name,

        /// <remarks/>
        outline,

        /// <remarks/>
        scheme,

        /// <remarks/>
        shadow,

        /// <remarks/>
        strike,

        /// <remarks/>
        sz,

        /// <remarks/>
        u,

        /// <remarks/>
        vertAlign,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Fills
    {

        private CT_Fill[] fillField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("fill")]
        public CT_Fill[] fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Fill
    {

        private object itemField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("gradientFill", typeof(CT_GradientFill))]
        [System.Xml.Serialization.XmlElementAttribute("patternFill", typeof(CT_PatternFill))]
        public object Item
        {
            get
            {
                return this.itemField;
            }
            set
            {
                this.itemField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_GradientFill
    {

        private CT_GradientStop1[] stopField;

        private ST_GradientType typeField;

        private double degreeField;

        private double leftField;

        private double rightField;

        private double topField;

        private double bottomField;

        public CT_GradientFill()
        {
            this.typeField = ST_GradientType.linear;
            this.degreeField = 0;
            this.leftField = 0;
            this.rightField = 0;
            this.topField = 0;
            this.bottomField = 0;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("stop")]
        public CT_GradientStop1[] stop
        {
            get
            {
                return this.stopField;
            }
            set
            {
                this.stopField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_GradientType.linear)]
        public ST_GradientType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double degree
        {
            get
            {
                return this.degreeField;
            }
            set
            {
                this.degreeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(TypeName = "CT_GradientStop", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_GradientStop1
    {

        private CT_Color1 colorField;

        private double positionField;

        /// <remarks/>
        public CT_Color1 color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double position
        {
            get
            {
                return this.positionField;
            }
            set
            {
                this.positionField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_GradientType
    {

        /// <remarks/>
        linear,

        /// <remarks/>
        path,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PatternFill
    {

        private CT_Color1 fgColorField;

        private CT_Color1 bgColorField;

        private ST_PatternType patternTypeField;

        private bool patternTypeFieldSpecified;

        /// <remarks/>
        public CT_Color1 fgColor
        {
            get
            {
                return this.fgColorField;
            }
            set
            {
                this.fgColorField = value;
            }
        }

        /// <remarks/>
        public CT_Color1 bgColor
        {
            get
            {
                return this.bgColorField;
            }
            set
            {
                this.bgColorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_PatternType patternType
        {
            get
            {
                return this.patternTypeField;
            }
            set
            {
                this.patternTypeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool patternTypeSpecified
        {
            get
            {
                return this.patternTypeFieldSpecified;
            }
            set
            {
                this.patternTypeFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PatternType
    {

        /// <remarks/>
        none,

        /// <remarks/>
        solid,

        /// <remarks/>
        mediumGray,

        /// <remarks/>
        darkGray,

        /// <remarks/>
        lightGray,

        /// <remarks/>
        darkHorizontal,

        /// <remarks/>
        darkVertical,

        /// <remarks/>
        darkDown,

        /// <remarks/>
        darkUp,

        /// <remarks/>
        darkGrid,

        /// <remarks/>
        darkTrellis,

        /// <remarks/>
        lightHorizontal,

        /// <remarks/>
        lightVertical,

        /// <remarks/>
        lightDown,

        /// <remarks/>
        lightUp,

        /// <remarks/>
        lightGrid,

        /// <remarks/>
        lightTrellis,

        /// <remarks/>
        gray125,

        /// <remarks/>
        gray0625,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Borders
    {

        private CT_Border[] borderField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("border")]
        public CT_Border[] border
        {
            get
            {
                return this.borderField;
            }
            set
            {
                this.borderField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Border
    {

        private CT_BorderPr leftField;

        private CT_BorderPr rightField;

        private CT_BorderPr topField;

        private CT_BorderPr bottomField;

        private CT_BorderPr diagonalField;

        private CT_BorderPr verticalField;

        private CT_BorderPr horizontalField;

        private bool diagonalUpField;

        private bool diagonalUpFieldSpecified;

        private bool diagonalDownField;

        private bool diagonalDownFieldSpecified;

        private bool outlineField;

        public CT_Border()
        {
            this.outlineField = true;
        }

        /// <remarks/>
        public CT_BorderPr left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        /// <remarks/>
        public CT_BorderPr right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        /// <remarks/>
        public CT_BorderPr top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        /// <remarks/>
        public CT_BorderPr bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        /// <remarks/>
        public CT_BorderPr diagonal
        {
            get
            {
                return this.diagonalField;
            }
            set
            {
                this.diagonalField = value;
            }
        }

        /// <remarks/>
        public CT_BorderPr vertical
        {
            get
            {
                return this.verticalField;
            }
            set
            {
                this.verticalField = value;
            }
        }

        /// <remarks/>
        public CT_BorderPr horizontal
        {
            get
            {
                return this.horizontalField;
            }
            set
            {
                this.horizontalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool diagonalUp
        {
            get
            {
                return this.diagonalUpField;
            }
            set
            {
                this.diagonalUpField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool diagonalUpSpecified
        {
            get
            {
                return this.diagonalUpFieldSpecified;
            }
            set
            {
                this.diagonalUpFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool diagonalDown
        {
            get
            {
                return this.diagonalDownField;
            }
            set
            {
                this.diagonalDownField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool diagonalDownSpecified
        {
            get
            {
                return this.diagonalDownFieldSpecified;
            }
            set
            {
                this.diagonalDownFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_BorderPr
    {

        private CT_Color1 colorField;

        private ST_BorderStyle styleField;

        public CT_BorderPr()
        {
            this.styleField = ST_BorderStyle.none;
        }

        /// <remarks/>
        public CT_Color1 color
        {
            get
            {
                return this.colorField;
            }
            set
            {
                this.colorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_BorderStyle.none)]
        public ST_BorderStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_BorderStyle
    {

        /// <remarks/>
        none,

        /// <remarks/>
        thin,

        /// <remarks/>
        medium,

        /// <remarks/>
        dashed,

        /// <remarks/>
        dotted,

        /// <remarks/>
        thick,

        /// <remarks/>
        @double,

        /// <remarks/>
        hair,

        /// <remarks/>
        mediumDashed,

        /// <remarks/>
        dashDot,

        /// <remarks/>
        mediumDashDot,

        /// <remarks/>
        dashDotDot,

        /// <remarks/>
        mediumDashDotDot,

        /// <remarks/>
        slantDashDot,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellStyleXfs
    {

        private CT_Xf[] xfField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("xf")]
        public CT_Xf[] xf
        {
            get
            {
                return this.xfField;
            }
            set
            {
                this.xfField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Xf
    {

        private CT_CellAlignment alignmentField;

        //private CT_CellProtection protectionField;

        //private CT_ExtensionList extLstField;

        private uint numFmtIdField;

        private bool numFmtIdFieldSpecified;

        private uint fontIdField;

        private bool fontIdFieldSpecified;

        private uint fillIdField;

        private bool fillIdFieldSpecified;

        private uint borderIdField;

        private bool borderIdFieldSpecified;

        private uint xfIdField;

        private bool xfIdFieldSpecified;

        private bool quotePrefixField;

        private bool pivotButtonField;

        private bool applyNumberFormatField;

        private bool applyNumberFormatFieldSpecified;

        private bool applyFontField;

        private bool applyFontFieldSpecified;

        private bool applyFillField;

        private bool applyFillFieldSpecified;

        private bool applyBorderField;

        private bool applyBorderFieldSpecified;

        private bool applyAlignmentField;

        private bool applyAlignmentFieldSpecified;

        private bool applyProtectionField;

        private bool applyProtectionFieldSpecified;

        public CT_Xf()
        {
            this.quotePrefixField = false;
            this.pivotButtonField = false;
        }

        /// <remarks/>
        public CT_CellAlignment alignment
        {
            get
            {
                return this.alignmentField;
            }
            set
            {
                this.alignmentField = value;
            }
        }

        /// <remarks/>
        //public CT_CellProtection protection
        //{
        //    get
        //    {
        //        return this.protectionField;
        //    }
        //    set
        //    {
        //        this.protectionField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_ExtensionList extLst
        //{
        //    get
        //    {
        //        return this.extLstField;
        //    }
        //    set
        //    {
        //        this.extLstField = value;
        //    }
        //}

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint numFmtId
        {
            get
            {
                return this.numFmtIdField;
            }
            set
            {
                this.numFmtIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool numFmtIdSpecified
        {
            get
            {
                return this.numFmtIdFieldSpecified;
            }
            set
            {
                this.numFmtIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fontId
        {
            get
            {
                return this.fontIdField;
            }
            set
            {
                this.fontIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fontIdSpecified
        {
            get
            {
                return this.fontIdFieldSpecified;
            }
            set
            {
                this.fontIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fillId
        {
            get
            {
                return this.fillIdField;
            }
            set
            {
                this.fillIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fillIdSpecified
        {
            get
            {
                return this.fillIdFieldSpecified;
            }
            set
            {
                this.fillIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint borderId
        {
            get
            {
                return this.borderIdField;
            }
            set
            {
                this.borderIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool borderIdSpecified
        {
            get
            {
                return this.borderIdFieldSpecified;
            }
            set
            {
                this.borderIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint xfId
        {
            get
            {
                return this.xfIdField;
            }
            set
            {
                this.xfIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool xfIdSpecified
        {
            get
            {
                return this.xfIdFieldSpecified;
            }
            set
            {
                this.xfIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool quotePrefix
        {
            get
            {
                return this.quotePrefixField;
            }
            set
            {
                this.quotePrefixField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool pivotButton
        {
            get
            {
                return this.pivotButtonField;
            }
            set
            {
                this.pivotButtonField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyNumberFormat
        {
            get
            {
                return this.applyNumberFormatField;
            }
            set
            {
                this.applyNumberFormatField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyNumberFormatSpecified
        {
            get
            {
                return this.applyNumberFormatFieldSpecified;
            }
            set
            {
                this.applyNumberFormatFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyFont
        {
            get
            {
                return this.applyFontField;
            }
            set
            {
                this.applyFontField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyFontSpecified
        {
            get
            {
                return this.applyFontFieldSpecified;
            }
            set
            {
                this.applyFontFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyFill
        {
            get
            {
                return this.applyFillField;
            }
            set
            {
                this.applyFillField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyFillSpecified
        {
            get
            {
                return this.applyFillFieldSpecified;
            }
            set
            {
                this.applyFillFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyBorder
        {
            get
            {
                return this.applyBorderField;
            }
            set
            {
                this.applyBorderField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyBorderSpecified
        {
            get
            {
                return this.applyBorderFieldSpecified;
            }
            set
            {
                this.applyBorderFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyAlignment
        {
            get
            {
                return this.applyAlignmentField;
            }
            set
            {
                this.applyAlignmentField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyAlignmentSpecified
        {
            get
            {
                return this.applyAlignmentFieldSpecified;
            }
            set
            {
                this.applyAlignmentFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool applyProtection
        {
            get
            {
                return this.applyProtectionField;
            }
            set
            {
                this.applyProtectionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool applyProtectionSpecified
        {
            get
            {
                return this.applyProtectionFieldSpecified;
            }
            set
            {
                this.applyProtectionFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellAlignment
    {

        private ST_HorizontalAlignment horizontalField;

        private bool horizontalFieldSpecified;

        private ST_VerticalAlignment verticalField;

        private bool verticalFieldSpecified;

        private uint textRotationField;

        private bool textRotationFieldSpecified;

        private bool wrapTextField;

        private bool wrapTextFieldSpecified;

        private uint indentField;

        private bool indentFieldSpecified;

        private int relativeIndentField;

        private bool relativeIndentFieldSpecified;

        private bool justifyLastLineField;

        private bool justifyLastLineFieldSpecified;

        private bool shrinkToFitField;

        private bool shrinkToFitFieldSpecified;

        private uint readingOrderField;

        private bool readingOrderFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_HorizontalAlignment horizontal
        {
            get
            {
                return this.horizontalField;
            }
            set
            {
                this.horizontalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool horizontalSpecified
        {
            get
            {
                return this.horizontalFieldSpecified;
            }
            set
            {
                this.horizontalFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_VerticalAlignment vertical
        {
            get
            {
                return this.verticalField;
            }
            set
            {
                this.verticalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool verticalSpecified
        {
            get
            {
                return this.verticalFieldSpecified;
            }
            set
            {
                this.verticalFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint textRotation
        {
            get
            {
                return this.textRotationField;
            }
            set
            {
                this.textRotationField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool textRotationSpecified
        {
            get
            {
                return this.textRotationFieldSpecified;
            }
            set
            {
                this.textRotationFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool wrapText
        {
            get
            {
                return this.wrapTextField;
            }
            set
            {
                this.wrapTextField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool wrapTextSpecified
        {
            get
            {
                return this.wrapTextFieldSpecified;
            }
            set
            {
                this.wrapTextFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint indent
        {
            get
            {
                return this.indentField;
            }
            set
            {
                this.indentField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool indentSpecified
        {
            get
            {
                return this.indentFieldSpecified;
            }
            set
            {
                this.indentFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int relativeIndent
        {
            get
            {
                return this.relativeIndentField;
            }
            set
            {
                this.relativeIndentField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool relativeIndentSpecified
        {
            get
            {
                return this.relativeIndentFieldSpecified;
            }
            set
            {
                this.relativeIndentFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool justifyLastLine
        {
            get
            {
                return this.justifyLastLineField;
            }
            set
            {
                this.justifyLastLineField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool justifyLastLineSpecified
        {
            get
            {
                return this.justifyLastLineFieldSpecified;
            }
            set
            {
                this.justifyLastLineFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool shrinkToFit
        {
            get
            {
                return this.shrinkToFitField;
            }
            set
            {
                this.shrinkToFitField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool shrinkToFitSpecified
        {
            get
            {
                return this.shrinkToFitFieldSpecified;
            }
            set
            {
                this.shrinkToFitFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint readingOrder
        {
            get
            {
                return this.readingOrderField;
            }
            set
            {
                this.readingOrderField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool readingOrderSpecified
        {
            get
            {
                return this.readingOrderFieldSpecified;
            }
            set
            {
                this.readingOrderFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_HorizontalAlignment
    {

        /// <remarks/>
        general,

        /// <remarks/>
        left,

        /// <remarks/>
        center,

        /// <remarks/>
        right,

        /// <remarks/>
        fill,

        /// <remarks/>
        justify,

        /// <remarks/>
        centerContinuous,

        /// <remarks/>
        distributed,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_VerticalAlignment
    {

        /// <remarks/>
        top,

        /// <remarks/>
        center,

        /// <remarks/>
        bottom,

        /// <remarks/>
        justify,

        /// <remarks/>
        distributed,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Colors
    {

        private CT_RgbColor[] indexedColorsField;

        private CT_Color1[] mruColorsField;

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("rgbColor", IsNullable = false)]
        public CT_RgbColor[] indexedColors
        {
            get
            {
                return this.indexedColorsField;
            }
            set
            {
                this.indexedColorsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("color", IsNullable = false)]
        public CT_Color1[] mruColors
        {
            get
            {
                return this.mruColorsField;
            }
            set
            {
                this.mruColorsField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_RgbColor
    {

        private byte[] rgbField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType = "hexBinary")]
        public byte[] rgb
        {
            get
            {
                return this.rgbField;
            }
            set
            {
                this.rgbField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_Sst
    {

        private CT_Rst[] siField;

        private CT_ExtensionList extLstField;

        private uint countField;

        private bool countFieldSpecified;

        private uint uniqueCountField;

        private bool uniqueCountFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("si")]
        public CT_Rst[] si
        {
            get
            {
                return this.siField;
            }
            set
            {
                this.siField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint uniqueCount
        {
            get
            {
                return this.uniqueCountField;
            }
            set
            {
                this.uniqueCountField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool uniqueCountSpecified
        {
            get
            {
                return this.uniqueCountFieldSpecified;
            }
            set
            {
                this.uniqueCountFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Rst
    {

        private string tField;

        private CT_RElt[] rField;

        private CT_PhoneticRun[] rPhField;

        private CT_PhoneticPr phoneticPrField;

        /// <remarks/>
        public string t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("r")]
        public CT_RElt[] r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("rPh")]
        public CT_PhoneticRun[] rPh
        {
            get
            {
                return this.rPhField;
            }
            set
            {
                this.rPhField = value;
            }
        }

        /// <remarks/>
        public CT_PhoneticPr phoneticPr
        {
            get
            {
                return this.phoneticPrField;
            }
            set
            {
                this.phoneticPrField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_RElt
    {

        private CT_RPrElt rPrField;

        private string tField;

        /// <remarks/>
        public CT_RPrElt rPr
        {
            get
            {
                return this.rPrField;
            }
            set
            {
                this.rPrField = value;
            }
        }

        /// <remarks/>
        public string t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_RPrElt
    {

        private object[] itemsField;

        private ItemsChoiceType4[] itemsElementNameField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("b", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("charset", typeof(CT_IntProperty))]
        [System.Xml.Serialization.XmlElementAttribute("color", typeof(CT_Color1))]
        [System.Xml.Serialization.XmlElementAttribute("condense", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("extend", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("family", typeof(CT_IntProperty))]
        [System.Xml.Serialization.XmlElementAttribute("i", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("outline", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("rFont", typeof(CT_FontName))]
        [System.Xml.Serialization.XmlElementAttribute("scheme", typeof(CT_FontScheme1))]
        [System.Xml.Serialization.XmlElementAttribute("shadow", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("strike", typeof(CT_BooleanProperty))]
        [System.Xml.Serialization.XmlElementAttribute("sz", typeof(CT_FontSize))]
        [System.Xml.Serialization.XmlElementAttribute("u", typeof(CT_UnderlineProperty))]
        [System.Xml.Serialization.XmlElementAttribute("vertAlign", typeof(CT_VerticalAlignFontProperty))]
        [System.Xml.Serialization.XmlChoiceIdentifierAttribute("ItemsElementName")]
        public object[] Items
        {
            get
            {
                return this.itemsField;
            }
            set
            {
                this.itemsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ItemsElementName")]
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public ItemsChoiceType4[] ItemsElementName
        {
            get
            {
                return this.itemsElementNameField;
            }
            set
            {
                this.itemsElementNameField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IncludeInSchema = false)]
    public enum ItemsChoiceType4
    {

        /// <remarks/>
        b,

        /// <remarks/>
        charset,

        /// <remarks/>
        color,

        /// <remarks/>
        condense,

        /// <remarks/>
        extend,

        /// <remarks/>
        family,

        /// <remarks/>
        i,

        /// <remarks/>
        outline,

        /// <remarks/>
        rFont,

        /// <remarks/>
        scheme,

        /// <remarks/>
        shadow,

        /// <remarks/>
        strike,

        /// <remarks/>
        sz,

        /// <remarks/>
        u,

        /// <remarks/>
        vertAlign,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PhoneticRun
    {

        private string tField;

        private uint sbField;

        private uint ebField;

        /// <remarks/>
        public string t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint sb
        {
            get
            {
                return this.sbField;
            }
            set
            {
                this.sbField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint eb
        {
            get
            {
                return this.ebField;
            }
            set
            {
                this.ebField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PhoneticPr
    {

        private uint fontIdField;

        private ST_PhoneticType typeField;

        private ST_PhoneticAlignment alignmentField;

        public CT_PhoneticPr()
        {
            this.typeField = ST_PhoneticType.fullwidthKatakana;
            this.alignmentField = ST_PhoneticAlignment.left;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fontId
        {
            get
            {
                return this.fontIdField;
            }
            set
            {
                this.fontIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_PhoneticType.fullwidthKatakana)]
        public ST_PhoneticType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_PhoneticAlignment.left)]
        public ST_PhoneticAlignment alignment
        {
            get
            {
                return this.alignmentField;
            }
            set
            {
                this.alignmentField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PhoneticType
    {

        /// <remarks/>
        halfwidthKatakana,

        /// <remarks/>
        fullwidthKatakana,

        /// <remarks/>
        Hiragana,

        /// <remarks/>
        noConversion,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PhoneticAlignment
    {

        /// <remarks/>
        noControl,

        /// <remarks/>
        left,

        /// <remarks/>
        center,

        /// <remarks/>
        distributed,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Extension
    {

        private System.Xml.XmlElement anyField;

        private string uriField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAnyElementAttribute()]
        public System.Xml.XmlElement Any
        {
            get
            {
                return this.anyField;
            }
            set
            {
                this.anyField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType = "token")]
        public string uri
        {
            get
            {
                return this.uriField;
            }
            set
            {
                this.uriField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellXfs
    {

        private CT_Xf[] xfField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("xf")]
        public CT_Xf[] xf
        {
            get
            {
                return this.xfField;
            }
            set
            {
                this.xfField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellStyles
    {

        private CT_CellStyle[] cellStyleField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("cellStyle")]
        public CT_CellStyle[] cellStyle
        {
            get
            {
                return this.cellStyleField;
            }
            set
            {
                this.cellStyleField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellStyle
    {

        private CT_ExtensionList extLstField;

        private string nameField;

        private uint xfIdField;

        private uint builtinIdField;

        private bool builtinIdFieldSpecified;

        private uint iLevelField;

        private bool iLevelFieldSpecified;

        private bool hiddenField;

        private bool hiddenFieldSpecified;

        private bool customBuiltinField;

        private bool customBuiltinFieldSpecified;

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint xfId
        {
            get
            {
                return this.xfIdField;
            }
            set
            {
                this.xfIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint builtinId
        {
            get
            {
                return this.builtinIdField;
            }
            set
            {
                this.builtinIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool builtinIdSpecified
        {
            get
            {
                return this.builtinIdFieldSpecified;
            }
            set
            {
                this.builtinIdFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint iLevel
        {
            get
            {
                return this.iLevelField;
            }
            set
            {
                this.iLevelField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool iLevelSpecified
        {
            get
            {
                return this.iLevelFieldSpecified;
            }
            set
            {
                this.iLevelFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hiddenSpecified
        {
            get
            {
                return this.hiddenFieldSpecified;
            }
            set
            {
                this.hiddenFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool customBuiltin
        {
            get
            {
                return this.customBuiltinField;
            }
            set
            {
                this.customBuiltinField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool customBuiltinSpecified
        {
            get
            {
                return this.customBuiltinFieldSpecified;
            }
            set
            {
                this.customBuiltinFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Dxfs
    {

        private CT_Dxf[] dxfField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("dxf")]
        public CT_Dxf[] dxf
        {
            get
            {
                return this.dxfField;
            }
            set
            {
                this.dxfField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Dxf
    {

        private CT_Font fontField;

        private CT_NumFmt numFmtField;

        private CT_Fill fillField;

        private CT_CellAlignment alignmentField;

        private CT_Border borderField;

        private CT_CellProtection protectionField;

        private CT_ExtensionList extLstField;

        /// <remarks/>
        public CT_Font font
        {
            get
            {
                return this.fontField;
            }
            set
            {
                this.fontField = value;
            }
        }

        /// <remarks/>
        public CT_NumFmt numFmt
        {
            get
            {
                return this.numFmtField;
            }
            set
            {
                this.numFmtField = value;
            }
        }

        /// <remarks/>
        public CT_Fill fill
        {
            get
            {
                return this.fillField;
            }
            set
            {
                this.fillField = value;
            }
        }

        /// <remarks/>
        public CT_CellAlignment alignment
        {
            get
            {
                return this.alignmentField;
            }
            set
            {
                this.alignmentField = value;
            }
        }

        /// <remarks/>
        public CT_Border border
        {
            get
            {
                return this.borderField;
            }
            set
            {
                this.borderField = value;
            }
        }

        /// <remarks/>
        public CT_CellProtection protection
        {
            get
            {
                return this.protectionField;
            }
            set
            {
                this.protectionField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_TableStyles
    {

        private CT_TableStyle[] tableStyleField;

        private uint countField;

        private bool countFieldSpecified;

        private string defaultTableStyleField;

        private string defaultPivotStyleField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("tableStyle")]
        public CT_TableStyle[] tableStyle
        {
            get
            {
                return this.tableStyleField;
            }
            set
            {
                this.tableStyleField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string defaultTableStyle
        {
            get
            {
                return this.defaultTableStyleField;
            }
            set
            {
                this.defaultTableStyleField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string defaultPivotStyle
        {
            get
            {
                return this.defaultPivotStyleField;
            }
            set
            {
                this.defaultPivotStyleField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_TableStyle
    {

        private CT_TableStyleElement[] tableStyleElementField;

        private string nameField;

        private bool pivotField;

        private bool tableField;

        private uint countField;

        private bool countFieldSpecified;

        public CT_TableStyle()
        {
            this.pivotField = true;
            this.tableField = true;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("tableStyleElement")]
        public CT_TableStyleElement[] tableStyleElement
        {
            get
            {
                return this.tableStyleElementField;
            }
            set
            {
                this.tableStyleElementField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool pivot
        {
            get
            {
                return this.pivotField;
            }
            set
            {
                this.pivotField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool table
        {
            get
            {
                return this.tableField;
            }
            set
            {
                this.tableField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_TableStyleElement
    {

        private ST_TableStyleType typeField;

        private uint sizeField;

        private uint dxfIdField;

        private bool dxfIdFieldSpecified;

        public CT_TableStyleElement()
        {
            this.sizeField = ((uint)(1));
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_TableStyleType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint size
        {
            get
            {
                return this.sizeField;
            }
            set
            {
                this.sizeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint dxfId
        {
            get
            {
                return this.dxfIdField;
            }
            set
            {
                this.dxfIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool dxfIdSpecified
        {
            get
            {
                return this.dxfIdFieldSpecified;
            }
            set
            {
                this.dxfIdFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_TableStyleType
    {

        /// <remarks/>
        wholeTable,

        /// <remarks/>
        headerRow,

        /// <remarks/>
        totalRow,

        /// <remarks/>
        firstColumn,

        /// <remarks/>
        lastColumn,

        /// <remarks/>
        firstRowStripe,

        /// <remarks/>
        secondRowStripe,

        /// <remarks/>
        firstColumnStripe,

        /// <remarks/>
        secondColumnStripe,

        /// <remarks/>
        firstHeaderCell,

        /// <remarks/>
        lastHeaderCell,

        /// <remarks/>
        firstTotalCell,

        /// <remarks/>
        lastTotalCell,

        /// <remarks/>
        firstSubtotalColumn,

        /// <remarks/>
        secondSubtotalColumn,

        /// <remarks/>
        thirdSubtotalColumn,

        /// <remarks/>
        firstSubtotalRow,

        /// <remarks/>
        secondSubtotalRow,

        /// <remarks/>
        thirdSubtotalRow,

        /// <remarks/>
        blankRow,

        /// <remarks/>
        firstColumnSubheading,

        /// <remarks/>
        secondColumnSubheading,

        /// <remarks/>
        thirdColumnSubheading,

        /// <remarks/>
        firstRowSubheading,

        /// <remarks/>
        secondRowSubheading,

        /// <remarks/>
        thirdRowSubheading,

        /// <remarks/>
        pageFieldLabels,

        /// <remarks/>
        pageFieldValues,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_ExtensionList
    {

        private CT_Extension[] extField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ext")]
        public CT_Extension[] ext
        {
            get
            {
                return this.extField;
            }
            set
            {
                this.extField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellProtection
    {

        private bool lockedField;

        private bool lockedFieldSpecified;

        private bool hiddenField;

        private bool hiddenFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool locked
        {
            get
            {
                return this.lockedField;
            }
            set
            {
                this.lockedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool lockedSpecified
        {
            get
            {
                return this.lockedFieldSpecified;
            }
            set
            {
                this.lockedFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool hiddenSpecified
        {
            get
            {
                return this.hiddenFieldSpecified;
            }
            set
            {
                this.hiddenFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("workbook", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_Workbook
    {

        //private CT_FileVersion fileVersionField;

        //private CT_FileSharing fileSharingField;

        //private CT_WorkbookPr workbookPrField;

        //private CT_WorkbookProtection workbookProtectionField;

        private CT_BookView[] bookViewsField;

        private CT_Sheet[] sheetsField;

        //private CT_FunctionGroups functionGroupsField;

        //private CT_ExternalReference[] externalReferencesField;

        //private CT_DefinedName[] definedNamesField;

        //private CT_CalcPr calcPrField;

        //private CT_OleSize oleSizeField;

        //private CT_CustomWorkbookView[] customWorkbookViewsField;

        //private CT_PivotCache[] pivotCachesField;

        //private CT_SmartTagPr smartTagPrField;

        //private CT_SmartTagType[] smartTagTypesField;

        //private CT_WebPublishing webPublishingField;

        //private CT_FileRecoveryPr[] fileRecoveryPrField;

        //private CT_WebPublishObjects webPublishObjectsField;

        //private CT_ExtensionList extLstField;

        /// <remarks/>
        //public CT_FileVersion fileVersion {
        //    get {
        //        return this.fileVersionField;
        //    }
        //    set {
        //        this.fileVersionField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_FileSharing fileSharing {
        //    get {
        //        return this.fileSharingField;
        //    }
        //    set {
        //        this.fileSharingField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_WorkbookPr workbookPr {
        //    get {
        //        return this.workbookPrField;
        //    }
        //    set {
        //        this.workbookPrField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_WorkbookProtection workbookProtection {
        //    get {
        //        return this.workbookProtectionField;
        //    }
        //    set {
        //        this.workbookProtectionField = value;
        //    }
        //}
        
        [System.Xml.Serialization.XmlArrayItemAttribute("workbookView", IsNullable=false)]
        public CT_BookView[] bookViews {
            get {
                return this.bookViewsField;
            }
            set {
                this.bookViewsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("sheet", IsNullable = false)]
        public CT_Sheet[] sheets
        {
            get
            {
                return this.sheetsField;
            }
            set
            {
                this.sheetsField = value;
            }
        }

        /// <remarks/>
        //public CT_FunctionGroups functionGroups {
        //    get {
        //        return this.functionGroupsField;
        //    }
        //    set {
        //        this.functionGroupsField = value;
        //    }
        //}

        /// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("externalReference", IsNullable=false)]
        //public CT_ExternalReference[] externalReferences {
        //    get {
        //        return this.externalReferencesField;
        //    }
        //    set {
        //        this.externalReferencesField = value;
        //    }
        //}

        /// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("definedName", IsNullable=false)]
        //public CT_DefinedName[] definedNames {
        //    get {
        //        return this.definedNamesField;
        //    }
        //    set {
        //        this.definedNamesField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_CalcPr calcPr {
        //    get {
        //        return this.calcPrField;
        //    }
        //    set {
        //        this.calcPrField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_OleSize oleSize {
        //    get {
        //        return this.oleSizeField;
        //    }
        //    set {
        //        this.oleSizeField = value;
        //    }
        //}
       
        /////<remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("customWorkbookView", IsNullable=false)]
        //public CT_CustomWorkbookView[] customWorkbookViews {
        //    get {
        //        return this.customWorkbookViewsField;
        //    }
        //    set {
        //        this.customWorkbookViewsField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("pivotCache", IsNullable=false)]
        //public CT_PivotCache[] pivotCaches {
        //    get {
        //        return this.pivotCachesField;
        //    }
        //    set {
        //        this.pivotCachesField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_SmartTagPr smartTagPr {
        //    get {
        //        return this.smartTagPrField;
        //    }
        //    set {
        //        this.smartTagPrField = value;
        //    }
        //}

        /// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("smartTagType", IsNullable=false)]
        //public CT_SmartTagType[] smartTagTypes {
        //    get {
        //        return this.smartTagTypesField;
        //    }
        //    set {
        //        this.smartTagTypesField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_WebPublishing webPublishing {
        //    get {
        //        return this.webPublishingField;
        //    }
        //    set {
        //        this.webPublishingField = value;
        //    }
        //}

        /// <remarks/>
        //[System.Xml.Serialization.XmlElementAttribute("fileRecoveryPr")]
        //public CT_FileRecoveryPr[] fileRecoveryPr {
        //    get {
        //        return this.fileRecoveryPrField;
        //    }
        //    set {
        //        this.fileRecoveryPrField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_WebPublishObjects webPublishObjects {
        //    get {
        //        return this.webPublishObjectsField;
        //    }
        //    set {
        //        this.webPublishObjectsField = value;
        //    }
        //}

        /// <remarks/>
        //public CT_ExtensionList extLst {
        //    get {
        //        return this.extLstField;
        //    }
        //    set {
        //        this.extLstField = value;
        //    }
        //}
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Sheet
    {

        private string nameField;

        private uint sheetIdField;

        private ST_SheetState stateField;

        private string idField;

        public CT_Sheet()
        {
            this.stateField = ST_SheetState.visible;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint sheetId
        {
            get
            {
                return this.sheetIdField;
            }
            set
            {
                this.sheetIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_SheetState.visible)]
        public ST_SheetState state
        {
            get
            {
                return this.stateField;
            }
            set
            {
                this.stateField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_SheetState
    {

        /// <remarks/>
        visible,

        /// <remarks/>
        hidden,

        /// <remarks/>
        veryHidden,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_CellFormulaType
    {

        /// <remarks/>
        normal,

        /// <remarks/>
        array,

        /// <remarks/>
        dataTable,

        /// <remarks/>
        shared,
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_CellType
    {

        /// <remarks/>
        b,

        /// <remarks/>
        n,

        /// <remarks/>
        e,

        /// <remarks/>
        s,

        /// <remarks/>
        str,

        /// <remarks/>
        inlineStr,
    }
    

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Row
    {

        private CT_Cell[] cField;

        private CT_ExtensionList extLstField;

        private uint rField;

        private bool rFieldSpecified;

        private string[] spansField;

        private uint sField;

        private bool customFormatField;

        private double htField;

        private bool htFieldSpecified;

        private bool hiddenField;

        private bool customHeightField;

        private byte outlineLevelField;

        private bool collapsedField;

        private bool thickTopField;

        private bool thickBotField;

        private bool phField;

        public CT_Row()
        {
            this.sField = ((uint)(0));
            this.customFormatField = false;
            this.hiddenField = false;
            this.customHeightField = false;
            this.outlineLevelField = ((byte)(0));
            this.collapsedField = false;
            this.thickTopField = false;
            this.thickBotField = false;
            this.phField = false;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("c")]
        public CT_Cell[] c
        {
            get
            {
                return this.cField;
            }
            set
            {
                this.cField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool rSpecified
        {
            get
            {
                return this.rFieldSpecified;
            }
            set
            {
                this.rFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string[] spans
        {
            get
            {
                return this.spansField;
            }
            set
            {
                this.spansField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint s
        {
            get
            {
                return this.sField;
            }
            set
            {
                this.sField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool customFormat
        {
            get
            {
                return this.customFormatField;
            }
            set
            {
                this.customFormatField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double ht
        {
            get
            {
                return this.htField;
            }
            set
            {
                this.htField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool htSpecified
        {
            get
            {
                return this.htFieldSpecified;
            }
            set
            {
                this.htFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool customHeight
        {
            get
            {
                return this.customHeightField;
            }
            set
            {
                this.customHeightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte outlineLevel
        {
            get
            {
                return this.outlineLevelField;
            }
            set
            {
                this.outlineLevelField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool collapsed
        {
            get
            {
                return this.collapsedField;
            }
            set
            {
                this.collapsedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool thickTop
        {
            get
            {
                return this.thickTopField;
            }
            set
            {
                this.thickTopField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool thickBot
        {
            get
            {
                return this.thickBotField;
            }
            set
            {
                this.thickBotField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool ph
        {
            get
            {
                return this.phField;
            }
            set
            {
                this.phField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Cell
    {

        private CT_CellFormula fField;

        private string vField;

        private CT_Rst isField;

        private CT_ExtensionList extLstField;

        private string rField;

        private uint sField;

        private ST_CellType tField;

        private uint cmField;

        private uint vmField;

        private bool phField;

        public CT_Cell()
        {
            this.sField = ((uint)(0));
            this.tField = ST_CellType.n;
            this.cmField = ((uint)(0));
            this.vmField = ((uint)(0));
            this.phField = false;
        }

        /// <remarks/>
        public CT_CellFormula f
        {
            get
            {
                return this.fField;
            }
            set
            {
                this.fField = value;
            }
        }

        /// <remarks/>
        public string v
        {
            get
            {
                return this.vField;
            }
            set
            {
                this.vField = value;
            }
        }

        /// <remarks/>
        public CT_Rst @is
        {
            get
            {
                return this.isField;
            }
            set
            {
                this.isField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string r
        {
            get
            {
                return this.rField;
            }
            set
            {
                this.rField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint s
        {
            get
            {
                return this.sField;
            }
            set
            {
                this.sField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_CellType.n)]
        public ST_CellType t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint cm
        {
            get
            {
                return this.cmField;
            }
            set
            {
                this.cmField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint vm
        {
            get
            {
                return this.vmField;
            }
            set
            {
                this.vmField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool ph
        {
            get
            {
                return this.phField;
            }
            set
            {
                this.phField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_CellFormula
    {

        private ST_CellFormulaType tField;

        private bool acaField;

        private string refField;

        private bool dt2DField;

        private bool dtrField;

        private bool del1Field;

        private bool del2Field;

        private string r1Field;

        private string r2Field;

        private bool caField;

        private uint siField;

        private bool siFieldSpecified;

        private bool bxField;

        private string valueField;

        public CT_CellFormula()
        {
            this.tField = ST_CellFormulaType.normal;
            this.acaField = false;
            this.dt2DField = false;
            this.dtrField = false;
            this.del1Field = false;
            this.del2Field = false;
            this.caField = false;
            this.bxField = false;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_CellFormulaType.normal)]
        public ST_CellFormulaType t
        {
            get
            {
                return this.tField;
            }
            set
            {
                this.tField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool aca
        {
            get
            {
                return this.acaField;
            }
            set
            {
                this.acaField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dt2D
        {
            get
            {
                return this.dt2DField;
            }
            set
            {
                this.dt2DField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool dtr
        {
            get
            {
                return this.dtrField;
            }
            set
            {
                this.dtrField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool del1
        {
            get
            {
                return this.del1Field;
            }
            set
            {
                this.del1Field = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool del2
        {
            get
            {
                return this.del2Field;
            }
            set
            {
                this.del2Field = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string r1
        {
            get
            {
                return this.r1Field;
            }
            set
            {
                this.r1Field = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string r2
        {
            get
            {
                return this.r2Field;
            }
            set
            {
                this.r2Field = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool ca
        {
            get
            {
                return this.caField;
            }
            set
            {
                this.caField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint si
        {
            get
            {
                return this.siField;
            }
            set
            {
                this.siField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool siSpecified
        {
            get
            {
                return this.siFieldSpecified;
            }
            set
            {
                this.siFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool bx
        {
            get
            {
                return this.bxField;
            }
            set
            {
                this.bxField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string Value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [System.Xml.Serialization.XmlRootAttribute("worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", IsNullable = false)]
    public partial class CT_Worksheet
    {

        //private CT_SheetPr sheetPrField;

        //private CT_SheetDimension dimensionField;

        private CT_SheetViews sheetViewsField;

        private CT_SheetFormatPr sheetFormatPrField;

        private CT_Col[] colsField;

        private CT_Row[] sheetDataField;

        //private CT_SheetCalcPr sheetCalcPrField;

        //private CT_SheetProtection sheetProtectionField;

        //private CT_ProtectedRange[] protectedRangesField;

        //private CT_Scenarios scenariosField;

        //private CT_AutoFilter autoFilterField;

        //private CT_SortState sortStateField;

        //private CT_DataConsolidate dataConsolidateField;

        //private CT_CustomSheetView[] customSheetViewsField;

        private CT_MergeCells mergeCellsField;

        //private CT_PhoneticPr phoneticPrField;

        //private CT_ConditionalFormatting[] conditionalFormattingField;

        //private CT_DataValidations dataValidationsField;

        //private CT_Hyperlink1[] hyperlinksField;

        //private CT_PrintOptions printOptionsField;

        private CT_PageMargins pageMarginsField;

        private CT_PageSetup pageSetupField;

        //private CT_HeaderFooter headerFooterField;

        private CT_PageBreak rowBreaksField;

        private CT_PageBreak colBreaksField;

        //private CT_CustomProperty[] customPropertiesField;

        //private CT_CellWatch[] cellWatchesField;

        //private CT_IgnoredErrors ignoredErrorsField;

        //private CT_CellSmartTags[] smartTagsField;

        //private CT_Drawing drawingField;

        //private CT_LegacyDrawing legacyDrawingField;

        //private CT_LegacyDrawing legacyDrawingHFField;

        //private CT_SheetBackgroundPicture pictureField;

        //private CT_OleObject[] oleObjectsField;

        //private CT_Control[] controlsField;

        //private CT_WebPublishItems webPublishItemsField;

        //private CT_TableParts tablePartsField;

        //private CT_ExtensionList extLstField;

        ///// <remarks/>
        //public CT_SheetPr sheetPr {
        //    get {
        //        return this.sheetPrField;
        //    }
        //    set {
        //        this.sheetPrField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_SheetDimension dimension {
        //    get {
        //        return this.dimensionField;
        //    }
        //    set {
        //        this.dimensionField = value;
        //    }
        //}

        /// <remarks/>
        public CT_SheetViews sheetViews
        {
            get
            {
                return this.sheetViewsField;
            }
            set
            {
                this.sheetViewsField = value;
            }
        }

        /// <remarks/>
        public CT_SheetFormatPr sheetFormatPr
        {
            get
            {
                return this.sheetFormatPrField;
            }
            set
            {
                this.sheetFormatPrField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("col", typeof(CT_Col), IsNullable = false)]
        public CT_Col[] cols
        {
            get
            {
                return this.colsField;
            }
            set
            {
                this.colsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("row", IsNullable = false)]
        public CT_Row[] sheetData
        {
            get
            {
                return this.sheetDataField;
            }
            set
            {
                this.sheetDataField = value;
            }
        }

        ///// <remarks/>
        //public CT_SheetCalcPr sheetCalcPr {
        //    get {
        //        return this.sheetCalcPrField;
        //    }
        //    set {
        //        this.sheetCalcPrField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_SheetProtection sheetProtection {
        //    get {
        //        return this.sheetProtectionField;
        //    }
        //    set {
        //        this.sheetProtectionField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("protectedRange", IsNullable=false)]
        //public CT_ProtectedRange[] protectedRanges {
        //    get {
        //        return this.protectedRangesField;
        //    }
        //    set {
        //        this.protectedRangesField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_Scenarios scenarios {
        //    get {
        //        return this.scenariosField;
        //    }
        //    set {
        //        this.scenariosField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_AutoFilter autoFilter {
        //    get {
        //        return this.autoFilterField;
        //    }
        //    set {
        //        this.autoFilterField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_SortState sortState {
        //    get {
        //        return this.sortStateField;
        //    }
        //    set {
        //        this.sortStateField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_DataConsolidate dataConsolidate {
        //    get {
        //        return this.dataConsolidateField;
        //    }
        //    set {
        //        this.dataConsolidateField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("customSheetView", IsNullable=false)]
        //public CT_CustomSheetView[] customSheetViews {
        //    get {
        //        return this.customSheetViewsField;
        //    }
        //    set {
        //        this.customSheetViewsField = value;
        //    }
        //}

        /// <remarks/>
        public CT_MergeCells mergeCells
        {
            get
            {
                return this.mergeCellsField;
            }
            set
            {
                this.mergeCellsField = value;
            }
        }

        ///// <remarks/>
        //public CT_PhoneticPr phoneticPr {
        //    get {
        //        return this.phoneticPrField;
        //    }
        //    set {
        //        this.phoneticPrField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlElementAttribute("conditionalFormatting")]
        //public CT_ConditionalFormatting[] conditionalFormatting {
        //    get {
        //        return this.conditionalFormattingField;
        //    }
        //    set {
        //        this.conditionalFormattingField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_DataValidations dataValidations {
        //    get {
        //        return this.dataValidationsField;
        //    }
        //    set {
        //        this.dataValidationsField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("hyperlink", IsNullable=false)]
        //public CT_Hyperlink1[] hyperlinks {
        //    get {
        //        return this.hyperlinksField;
        //    }
        //    set {
        //        this.hyperlinksField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_PrintOptions printOptions {
        //    get {
        //        return this.printOptionsField;
        //    }
        //    set {
        //        this.printOptionsField = value;
        //    }
        //}

        /// <remarks/>
        public CT_PageMargins pageMargins
        {
            get
            {
                return this.pageMarginsField;
            }
            set
            {
                this.pageMarginsField = value;
            }
        }

        /// <remarks/>
        public CT_PageSetup pageSetup
        {
            get
            {
                return this.pageSetupField;
            }
            set
            {
                this.pageSetupField = value;
            }
        }

        ///// <remarks/>
        //public CT_HeaderFooter headerFooter {
        //    get {
        //        return this.headerFooterField;
        //    }
        //    set {
        //        this.headerFooterField = value;
        //    }
        //}

        /// <remarks/>
        public CT_PageBreak rowBreaks
        {
            get
            {
                return this.rowBreaksField;
            }
            set
            {
                this.rowBreaksField = value;
            }
        }

        /// <remarks/>
        public CT_PageBreak colBreaks
        {
            get
            {
                return this.colBreaksField;
            }
            set
            {
                this.colBreaksField = value;
            }
        }

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("customPr", IsNullable=false)]
        //public CT_CustomProperty[] customProperties {
        //    get {
        //        return this.customPropertiesField;
        //    }
        //    set {
        //        this.customPropertiesField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("cellWatch", IsNullable=false)]
        //public CT_CellWatch[] cellWatches {
        //    get {
        //        return this.cellWatchesField;
        //    }
        //    set {
        //        this.cellWatchesField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_IgnoredErrors ignoredErrors {
        //    get {
        //        return this.ignoredErrorsField;
        //    }
        //    set {
        //        this.ignoredErrorsField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("cellSmartTags", IsNullable=false)]
        //public CT_CellSmartTags[] smartTags {
        //    get {
        //        return this.smartTagsField;
        //    }
        //    set {
        //        this.smartTagsField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_Drawing drawing {
        //    get {
        //        return this.drawingField;
        //    }
        //    set {
        //        this.drawingField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_LegacyDrawing legacyDrawing {
        //    get {
        //        return this.legacyDrawingField;
        //    }
        //    set {
        //        this.legacyDrawingField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_LegacyDrawing legacyDrawingHF {
        //    get {
        //        return this.legacyDrawingHFField;
        //    }
        //    set {
        //        this.legacyDrawingHFField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_SheetBackgroundPicture picture {
        //    get {
        //        return this.pictureField;
        //    }
        //    set {
        //        this.pictureField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("oleObject", IsNullable=false)]
        //public CT_OleObject[] oleObjects {
        //    get {
        //        return this.oleObjectsField;
        //    }
        //    set {
        //        this.oleObjectsField = value;
        //    }
        //}

        ///// <remarks/>
        //[System.Xml.Serialization.XmlArrayItemAttribute("control", IsNullable=false)]
        //public CT_Control[] controls {
        //    get {
        //        return this.controlsField;
        //    }
        //    set {
        //        this.controlsField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_WebPublishItems webPublishItems {
        //    get {
        //        return this.webPublishItemsField;
        //    }
        //    set {
        //        this.webPublishItemsField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_TableParts tableParts {
        //    get {
        //        return this.tablePartsField;
        //    }
        //    set {
        //        this.tablePartsField = value;
        //    }
        //}

        ///// <remarks/>
        //public CT_ExtensionList extLst {
        //    get {
        //        return this.extLstField;
        //    }
        //    set {
        //        this.extLstField = value;
        //    }
        //}
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Col
    {

        private uint minField;

        private uint maxField;

        private double widthField;

        private bool widthFieldSpecified;

        private uint styleField;

        private bool hiddenField;

        private bool bestFitField;

        private bool customWidthField;

        private bool phoneticField;

        private byte outlineLevelField;

        private bool collapsedField;

        public CT_Col()
        {
            this.styleField = ((uint)(0));
            this.hiddenField = false;
            this.bestFitField = false;
            this.customWidthField = false;
            this.phoneticField = false;
            this.outlineLevelField = ((byte)(0));
            this.collapsedField = false;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint min
        {
            get
            {
                return this.minField;
            }
            set
            {
                this.minField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint max
        {
            get
            {
                return this.maxField;
            }
            set
            {
                this.maxField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double width
        {
            get
            {
                return this.widthField;
            }
            set
            {
                this.widthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool widthSpecified
        {
            get
            {
                return this.widthFieldSpecified;
            }
            set
            {
                this.widthFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool hidden
        {
            get
            {
                return this.hiddenField;
            }
            set
            {
                this.hiddenField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool bestFit
        {
            get
            {
                return this.bestFitField;
            }
            set
            {
                this.bestFitField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool customWidth
        {
            get
            {
                return this.customWidthField;
            }
            set
            {
                this.customWidthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool phonetic
        {
            get
            {
                return this.phoneticField;
            }
            set
            {
                this.phoneticField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte outlineLevel
        {
            get
            {
                return this.outlineLevelField;
            }
            set
            {
                this.outlineLevelField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool collapsed
        {
            get
            {
                return this.collapsedField;
            }
            set
            {
                this.collapsedField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_MergeCells
    {

        private CT_MergeCell[] mergeCellField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("mergeCell")]
        public CT_MergeCell[] mergeCell
        {
            get
            {
                return this.mergeCellField;
            }
            set
            {
                this.mergeCellField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_MergeCell
    {

        private string refField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string @ref
        {
            get
            {
                return this.refField;
            }
            set
            {
                this.refField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_SheetViews
    {

        private CT_SheetView[] sheetViewField;

        private CT_ExtensionList extLstField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("sheetView")]
        public CT_SheetView[] sheetView
        {
            get
            {
                return this.sheetViewField;
            }
            set
            {
                this.sheetViewField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "2.0.50727.3038")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_SheetView
    {

        private CT_Pane paneField;

        private CT_Selection[] selectionField;

        private CT_PivotSelection[] pivotSelectionField;

        private CT_ExtensionList extLstField;

        private bool windowProtectionField;

        private bool showFormulasField;

        private bool showGridLinesField;

        private bool showRowColHeadersField;

        private bool showZerosField;

        private bool rightToLeftField;

        private bool tabSelectedField;

        private bool showRulerField;

        private bool showOutlineSymbolsField;

        private bool defaultGridColorField;

        private bool showWhiteSpaceField;

        private ST_SheetViewType viewField;

        private string topLeftCellField;

        private uint colorIdField;

        private uint zoomScaleField;

        private uint zoomScaleNormalField;

        private uint zoomScaleSheetLayoutViewField;

        private uint zoomScalePageLayoutViewField;

        private uint workbookViewIdField;

        public CT_SheetView()
        {
            this.windowProtectionField = false;
            this.showFormulasField = false;
            this.showGridLinesField = true;
            this.showRowColHeadersField = true;
            this.showZerosField = true;
            this.rightToLeftField = false;
            this.tabSelectedField = false;
            this.showRulerField = true;
            this.showOutlineSymbolsField = true;
            this.defaultGridColorField = true;
            this.showWhiteSpaceField = true;
            this.viewField = ST_SheetViewType.normal;
            this.colorIdField = ((uint)(64));
            this.zoomScaleField = ((uint)(100));
            this.zoomScaleNormalField = ((uint)(0));
            this.zoomScaleSheetLayoutViewField = ((uint)(0));
            this.zoomScalePageLayoutViewField = ((uint)(0));
        }

        /// <remarks/>
        public CT_Pane pane
        {
            get
            {
                return this.paneField;
            }
            set
            {
                this.paneField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("selection")]
        public CT_Selection[] selection
        {
            get
            {
                return this.selectionField;
            }
            set
            {
                this.selectionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("pivotSelection")]
        public CT_PivotSelection[] pivotSelection
        {
            get
            {
                return this.pivotSelectionField;
            }
            set
            {
                this.pivotSelectionField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool windowProtection
        {
            get
            {
                return this.windowProtectionField;
            }
            set
            {
                this.windowProtectionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showFormulas
        {
            get
            {
                return this.showFormulasField;
            }
            set
            {
                this.showFormulasField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showGridLines
        {
            get
            {
                return this.showGridLinesField;
            }
            set
            {
                this.showGridLinesField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showRowColHeaders
        {
            get
            {
                return this.showRowColHeadersField;
            }
            set
            {
                this.showRowColHeadersField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showZeros
        {
            get
            {
                return this.showZerosField;
            }
            set
            {
                this.showZerosField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool rightToLeft
        {
            get
            {
                return this.rightToLeftField;
            }
            set
            {
                this.rightToLeftField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool tabSelected
        {
            get
            {
                return this.tabSelectedField;
            }
            set
            {
                this.tabSelectedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showRuler
        {
            get
            {
                return this.showRulerField;
            }
            set
            {
                this.showRulerField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showOutlineSymbols
        {
            get
            {
                return this.showOutlineSymbolsField;
            }
            set
            {
                this.showOutlineSymbolsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool defaultGridColor
        {
            get
            {
                return this.defaultGridColorField;
            }
            set
            {
                this.defaultGridColorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showWhiteSpace
        {
            get
            {
                return this.showWhiteSpaceField;
            }
            set
            {
                this.showWhiteSpaceField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_SheetViewType.normal)]
        public ST_SheetViewType view
        {
            get
            {
                return this.viewField;
            }
            set
            {
                this.viewField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string topLeftCell
        {
            get
            {
                return this.topLeftCellField;
            }
            set
            {
                this.topLeftCellField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "64")]
        public uint colorId
        {
            get
            {
                return this.colorIdField;
            }
            set
            {
                this.colorIdField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "100")]
        public uint zoomScale
        {
            get
            {
                return this.zoomScaleField;
            }
            set
            {
                this.zoomScaleField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint zoomScaleNormal
        {
            get
            {
                return this.zoomScaleNormalField;
            }
            set
            {
                this.zoomScaleNormalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint zoomScaleSheetLayoutView
        {
            get
            {
                return this.zoomScaleSheetLayoutViewField;
            }
            set
            {
                this.zoomScaleSheetLayoutViewField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint zoomScalePageLayoutView
        {
            get
            {
                return this.zoomScalePageLayoutViewField;
            }
            set
            {
                this.zoomScalePageLayoutViewField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint workbookViewId
        {
            get
            {
                return this.workbookViewIdField;
            }
            set
            {
                this.workbookViewIdField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Pane
    {

        private double xSplitField;

        private double ySplitField;

        private string topLeftCellField;

        private ST_Pane activePaneField;

        private ST_PaneState stateField;

        public CT_Pane()
        {
            this.xSplitField = 0;
            this.ySplitField = 0;
            this.activePaneField = ST_Pane.topLeft;
            this.stateField = ST_PaneState.split;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double xSplit
        {
            get
            {
                return this.xSplitField;
            }
            set
            {
                this.xSplitField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(0)]
        public double ySplit
        {
            get
            {
                return this.ySplitField;
            }
            set
            {
                this.ySplitField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string topLeftCell
        {
            get
            {
                return this.topLeftCellField;
            }
            set
            {
                this.topLeftCellField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Pane.topLeft)]
        public ST_Pane activePane
        {
            get
            {
                return this.activePaneField;
            }
            set
            {
                this.activePaneField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_PaneState.split)]
        public ST_PaneState state
        {
            get
            {
                return this.stateField;
            }
            set
            {
                this.stateField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_Pane
    {

        /// <remarks/>
        bottomRight,

        /// <remarks/>
        topRight,

        /// <remarks/>
        bottomLeft,

        /// <remarks/>
        topLeft,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PaneState
    {

        /// <remarks/>
        split,

        /// <remarks/>
        frozen,

        /// <remarks/>
        frozenSplit,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PivotSelection
    {

        private CT_PivotArea pivotAreaField;

        private ST_Pane paneField;

        private bool showHeaderField;

        private bool labelField;

        private bool dataField;

        private bool extendableField;

        private uint countField;

        private ST_Axis axisField;

        private bool axisFieldSpecified;

        private uint dimensionField;

        private uint startField;

        private uint minField;

        private uint maxField;

        private uint activeRowField;

        private uint activeColField;

        private uint previousRowField;

        private uint previousColField;

        private uint clickField;

        private string idField;

        public CT_PivotSelection()
        {
            this.paneField = ST_Pane.topLeft;
            this.showHeaderField = false;
            this.labelField = false;
            this.dataField = false;
            this.extendableField = false;
            this.countField = ((uint)(0));
            this.dimensionField = ((uint)(0));
            this.startField = ((uint)(0));
            this.minField = ((uint)(0));
            this.maxField = ((uint)(0));
            this.activeRowField = ((uint)(0));
            this.activeColField = ((uint)(0));
            this.previousRowField = ((uint)(0));
            this.previousColField = ((uint)(0));
            this.clickField = ((uint)(0));
        }

        /// <remarks/>
        public CT_PivotArea pivotArea
        {
            get
            {
                return this.pivotAreaField;
            }
            set
            {
                this.pivotAreaField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Pane.topLeft)]
        public ST_Pane pane
        {
            get
            {
                return this.paneField;
            }
            set
            {
                this.paneField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool showHeader
        {
            get
            {
                return this.showHeaderField;
            }
            set
            {
                this.showHeaderField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool label
        {
            get
            {
                return this.labelField;
            }
            set
            {
                this.labelField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool data
        {
            get
            {
                return this.dataField;
            }
            set
            {
                this.dataField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool extendable
        {
            get
            {
                return this.extendableField;
            }
            set
            {
                this.extendableField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_Axis axis
        {
            get
            {
                return this.axisField;
            }
            set
            {
                this.axisField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool axisSpecified
        {
            get
            {
                return this.axisFieldSpecified;
            }
            set
            {
                this.axisFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint dimension
        {
            get
            {
                return this.dimensionField;
            }
            set
            {
                this.dimensionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint start
        {
            get
            {
                return this.startField;
            }
            set
            {
                this.startField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint min
        {
            get
            {
                return this.minField;
            }
            set
            {
                this.minField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint max
        {
            get
            {
                return this.maxField;
            }
            set
            {
                this.maxField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint activeRow
        {
            get
            {
                return this.activeRowField;
            }
            set
            {
                this.activeRowField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint activeCol
        {
            get
            {
                return this.activeColField;
            }
            set
            {
                this.activeColField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint previousRow
        {
            get
            {
                return this.previousRowField;
            }
            set
            {
                this.previousRowField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint previousCol
        {
            get
            {
                return this.previousColField;
            }
            set
            {
                this.previousColField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint click
        {
            get
            {
                return this.clickField;
            }
            set
            {
                this.clickField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_SheetViewType
    {

        /// <remarks/>
        normal,

        /// <remarks/>
        pageBreakPreview,

        /// <remarks/>
        pageLayout,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Selection
    {

        private ST_Pane paneField;

        private string activeCellField;

        private uint activeCellIdField;

        private string[] sqrefField;

        public CT_Selection()
        {
            this.paneField = ST_Pane.topLeft;
            this.activeCellIdField = ((uint)(0));
            this.sqrefField = new string[] {
                    "A1"};
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Pane.topLeft)]
        public ST_Pane pane
        {
            get
            {
                return this.paneField;
            }
            set
            {
                this.paneField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string activeCell
        {
            get
            {
                return this.activeCellField;
            }
            set
            {
                this.activeCellField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint activeCellId
        {
            get
            {
                return this.activeCellIdField;
            }
            set
            {
                this.activeCellIdField = value;
            }
        }

        /// <remarks/>
        // CODEGEN Warning: DefaultValue attribute on members of type System.String[] is not supported in this version of the .Net Framework.
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string[] sqref
        {
            get
            {
                return this.sqrefField;
            }
            set
            {
                this.sqrefField = value;
            }
        }
    }

    // <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PivotArea
    {

        private CT_PivotAreaReferences referencesField;

        private CT_ExtensionList extLstField;

        private int fieldField;

        private bool fieldFieldSpecified;

        private ST_PivotAreaType typeField;

        private bool dataOnlyField;

        private bool labelOnlyField;

        private bool grandRowField;

        private bool grandColField;

        private bool cacheIndexField;

        private bool outlineField;

        private string offsetField;

        private bool collapsedLevelsAreSubtotalsField;

        private ST_Axis axisField;

        private bool axisFieldSpecified;

        private uint fieldPositionField;

        private bool fieldPositionFieldSpecified;

        public CT_PivotArea()
        {
            this.typeField = ST_PivotAreaType.normal;
            this.dataOnlyField = true;
            this.labelOnlyField = false;
            this.grandRowField = false;
            this.grandColField = false;
            this.cacheIndexField = false;
            this.outlineField = true;
            this.collapsedLevelsAreSubtotalsField = false;
        }

        /// <remarks/>
        public CT_PivotAreaReferences references
        {
            get
            {
                return this.referencesField;
            }
            set
            {
                this.referencesField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fieldSpecified
        {
            get
            {
                return this.fieldFieldSpecified;
            }
            set
            {
                this.fieldFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_PivotAreaType.normal)]
        public ST_PivotAreaType type
        {
            get
            {
                return this.typeField;
            }
            set
            {
                this.typeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool dataOnly
        {
            get
            {
                return this.dataOnlyField;
            }
            set
            {
                this.dataOnlyField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool labelOnly
        {
            get
            {
                return this.labelOnlyField;
            }
            set
            {
                this.labelOnlyField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool grandRow
        {
            get
            {
                return this.grandRowField;
            }
            set
            {
                this.grandRowField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool grandCol
        {
            get
            {
                return this.grandColField;
            }
            set
            {
                this.grandColField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool cacheIndex
        {
            get
            {
                return this.cacheIndexField;
            }
            set
            {
                this.cacheIndexField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool outline
        {
            get
            {
                return this.outlineField;
            }
            set
            {
                this.outlineField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string offset
        {
            get
            {
                return this.offsetField;
            }
            set
            {
                this.offsetField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool collapsedLevelsAreSubtotals
        {
            get
            {
                return this.collapsedLevelsAreSubtotalsField;
            }
            set
            {
                this.collapsedLevelsAreSubtotalsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ST_Axis axis
        {
            get
            {
                return this.axisField;
            }
            set
            {
                this.axisField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool axisSpecified
        {
            get
            {
                return this.axisFieldSpecified;
            }
            set
            {
                this.axisFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint fieldPosition
        {
            get
            {
                return this.fieldPositionField;
            }
            set
            {
                this.fieldPositionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fieldPositionSpecified
        {
            get
            {
                return this.fieldPositionFieldSpecified;
            }
            set
            {
                this.fieldPositionFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PivotAreaReferences
    {

        private CT_PivotAreaReference[] referenceField;

        private uint countField;

        private bool countFieldSpecified;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("reference")]
        public CT_PivotAreaReference[] reference
        {
            get
            {
                return this.referenceField;
            }
            set
            {
                this.referenceField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PivotAreaReference
    {

        private CT_Index[] xField;

        private CT_ExtensionList extLstField;

        private uint fieldField;

        private bool fieldFieldSpecified;

        private uint countField;

        private bool countFieldSpecified;

        private bool selectedField;

        private bool byPositionField;

        private bool relativeField;

        private bool defaultSubtotalField;

        private bool sumSubtotalField;

        private bool countASubtotalField;

        private bool avgSubtotalField;

        private bool maxSubtotalField;

        private bool minSubtotalField;

        private bool productSubtotalField;

        private bool countSubtotalField;

        private bool stdDevSubtotalField;

        private bool stdDevPSubtotalField;

        private bool varSubtotalField;

        private bool varPSubtotalField;

        public CT_PivotAreaReference()
        {
            this.selectedField = true;
            this.byPositionField = false;
            this.relativeField = false;
            this.defaultSubtotalField = false;
            this.sumSubtotalField = false;
            this.countASubtotalField = false;
            this.avgSubtotalField = false;
            this.maxSubtotalField = false;
            this.minSubtotalField = false;
            this.productSubtotalField = false;
            this.countSubtotalField = false;
            this.stdDevSubtotalField = false;
            this.stdDevPSubtotalField = false;
            this.varSubtotalField = false;
            this.varPSubtotalField = false;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("x")]
        public CT_Index[] x
        {
            get
            {
                return this.xField;
            }
            set
            {
                this.xField = value;
            }
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint field
        {
            get
            {
                return this.fieldField;
            }
            set
            {
                this.fieldField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool fieldSpecified
        {
            get
            {
                return this.fieldFieldSpecified;
            }
            set
            {
                this.fieldFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool countSpecified
        {
            get
            {
                return this.countFieldSpecified;
            }
            set
            {
                this.countFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool selected
        {
            get
            {
                return this.selectedField;
            }
            set
            {
                this.selectedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool byPosition
        {
            get
            {
                return this.byPositionField;
            }
            set
            {
                this.byPositionField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool relative
        {
            get
            {
                return this.relativeField;
            }
            set
            {
                this.relativeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool defaultSubtotal
        {
            get
            {
                return this.defaultSubtotalField;
            }
            set
            {
                this.defaultSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool sumSubtotal
        {
            get
            {
                return this.sumSubtotalField;
            }
            set
            {
                this.sumSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countASubtotal
        {
            get
            {
                return this.countASubtotalField;
            }
            set
            {
                this.countASubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool avgSubtotal
        {
            get
            {
                return this.avgSubtotalField;
            }
            set
            {
                this.avgSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool maxSubtotal
        {
            get
            {
                return this.maxSubtotalField;
            }
            set
            {
                this.maxSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool minSubtotal
        {
            get
            {
                return this.minSubtotalField;
            }
            set
            {
                this.minSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool productSubtotal
        {
            get
            {
                return this.productSubtotalField;
            }
            set
            {
                this.productSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool countSubtotal
        {
            get
            {
                return this.countSubtotalField;
            }
            set
            {
                this.countSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevSubtotal
        {
            get
            {
                return this.stdDevSubtotalField;
            }
            set
            {
                this.stdDevSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool stdDevPSubtotal
        {
            get
            {
                return this.stdDevPSubtotalField;
            }
            set
            {
                this.stdDevPSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varSubtotal
        {
            get
            {
                return this.varSubtotalField;
            }
            set
            {
                this.varSubtotalField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool varPSubtotal
        {
            get
            {
                return this.varPSubtotalField;
            }
            set
            {
                this.varPSubtotalField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PivotAreaType
    {

        /// <remarks/>
        none,

        /// <remarks/>
        normal,

        /// <remarks/>
        data,

        /// <remarks/>
        all,

        /// <remarks/>
        origin,

        /// <remarks/>
        button,

        /// <remarks/>
        topRight,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_Axis
    {

        /// <remarks/>
        axisRow,

        /// <remarks/>
        axisCol,

        /// <remarks/>
        axisPage,

        /// <remarks/>
        axisValues,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Index
    {

        private uint vField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint v
        {
            get
            {
                return this.vField;
            }
            set
            {
                this.vField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_SheetFormatPr
    {

        private uint baseColWidthField;

        private double defaultColWidthField;

        private bool defaultColWidthFieldSpecified;

        private double defaultRowHeightField;

        private bool customHeightField;

        private bool zeroHeightField;

        private bool thickTopField;

        private bool thickBottomField;

        private byte outlineLevelRowField;

        private byte outlineLevelColField;

        public CT_SheetFormatPr()
        {
            this.baseColWidthField = ((uint)(8));
            this.customHeightField = false;
            this.zeroHeightField = false;
            this.thickTopField = false;
            this.thickBottomField = false;
            this.outlineLevelRowField = ((byte)(0));
            this.outlineLevelColField = ((byte)(0));
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "8")]
        public uint baseColWidth
        {
            get
            {
                return this.baseColWidthField;
            }
            set
            {
                this.baseColWidthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double defaultColWidth
        {
            get
            {
                return this.defaultColWidthField;
            }
            set
            {
                this.defaultColWidthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool defaultColWidthSpecified
        {
            get
            {
                return this.defaultColWidthFieldSpecified;
            }
            set
            {
                this.defaultColWidthFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double defaultRowHeight
        {
            get
            {
                return this.defaultRowHeightField;
            }
            set
            {
                this.defaultRowHeightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool customHeight
        {
            get
            {
                return this.customHeightField;
            }
            set
            {
                this.customHeightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool zeroHeight
        {
            get
            {
                return this.zeroHeightField;
            }
            set
            {
                this.zeroHeightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool thickTop
        {
            get
            {
                return this.thickTopField;
            }
            set
            {
                this.thickTopField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool thickBottom
        {
            get
            {
                return this.thickBottomField;
            }
            set
            {
                this.thickBottomField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte outlineLevelRow
        {
            get
            {
                return this.outlineLevelRowField;
            }
            set
            {
                this.outlineLevelRowField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(byte), "0")]
        public byte outlineLevelCol
        {
            get
            {
                return this.outlineLevelColField;
            }
            set
            {
                this.outlineLevelColField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_BookView
    {

        private CT_ExtensionList extLstField;

        private ST_Visibility visibilityField;

        private bool minimizedField;

        private bool showHorizontalScrollField;

        private bool showVerticalScrollField;

        private bool showSheetTabsField;

        private int xWindowField;

        private bool xWindowFieldSpecified;

        private int yWindowField;

        private bool yWindowFieldSpecified;

        private uint windowWidthField;

        private bool windowWidthFieldSpecified;

        private uint windowHeightField;

        private bool windowHeightFieldSpecified;

        private uint tabRatioField;

        private uint firstSheetField;

        private uint activeTabField;

        private bool autoFilterDateGroupingField;

        public CT_BookView()
        {
            this.visibilityField = ST_Visibility.visible;
            this.minimizedField = false;
            this.showHorizontalScrollField = true;
            this.showVerticalScrollField = true;
            this.showSheetTabsField = true;
            this.tabRatioField = ((uint)(600));
            this.firstSheetField = ((uint)(0));
            this.activeTabField = ((uint)(0));
            this.autoFilterDateGroupingField = true;
        }

        /// <remarks/>
        public CT_ExtensionList extLst
        {
            get
            {
                return this.extLstField;
            }
            set
            {
                this.extLstField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Visibility.visible)]
        public ST_Visibility visibility
        {
            get
            {
                return this.visibilityField;
            }
            set
            {
                this.visibilityField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool minimized
        {
            get
            {
                return this.minimizedField;
            }
            set
            {
                this.minimizedField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showHorizontalScroll
        {
            get
            {
                return this.showHorizontalScrollField;
            }
            set
            {
                this.showHorizontalScrollField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showVerticalScroll
        {
            get
            {
                return this.showVerticalScrollField;
            }
            set
            {
                this.showVerticalScrollField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool showSheetTabs
        {
            get
            {
                return this.showSheetTabsField;
            }
            set
            {
                this.showSheetTabsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int xWindow
        {
            get
            {
                return this.xWindowField;
            }
            set
            {
                this.xWindowField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool xWindowSpecified
        {
            get
            {
                return this.xWindowFieldSpecified;
            }
            set
            {
                this.xWindowFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int yWindow
        {
            get
            {
                return this.yWindowField;
            }
            set
            {
                this.yWindowField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool yWindowSpecified
        {
            get
            {
                return this.yWindowFieldSpecified;
            }
            set
            {
                this.yWindowFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint windowWidth
        {
            get
            {
                return this.windowWidthField;
            }
            set
            {
                this.windowWidthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool windowWidthSpecified
        {
            get
            {
                return this.windowWidthFieldSpecified;
            }
            set
            {
                this.windowWidthFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public uint windowHeight
        {
            get
            {
                return this.windowHeightField;
            }
            set
            {
                this.windowHeightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool windowHeightSpecified
        {
            get
            {
                return this.windowHeightFieldSpecified;
            }
            set
            {
                this.windowHeightFieldSpecified = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "600")]
        public uint tabRatio
        {
            get
            {
                return this.tabRatioField;
            }
            set
            {
                this.tabRatioField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint firstSheet
        {
            get
            {
                return this.firstSheetField;
            }
            set
            {
                this.firstSheetField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint activeTab
        {
            get
            {
                return this.activeTabField;
            }
            set
            {
                this.activeTabField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool autoFilterDateGrouping
        {
            get
            {
                return this.autoFilterDateGroupingField;
            }
            set
            {
                this.autoFilterDateGroupingField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_Visibility
    {

        /// <remarks/>
        visible,

        /// <remarks/>
        hidden,

        /// <remarks/>
        veryHidden,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PageBreak
    {

        private CT_Break[] brkField;

        private uint countField;

        private uint manualBreakCountField;

        public CT_PageBreak()
        {
            this.countField = ((uint)(0));
            this.manualBreakCountField = ((uint)(0));
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("brk")]
        public CT_Break[] brk
        {
            get
            {
                return this.brkField;
            }
            set
            {
                this.brkField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint manualBreakCount
        {
            get
            {
                return this.manualBreakCountField;
            }
            set
            {
                this.manualBreakCountField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_Break
    {

        private uint idField;

        private uint minField;

        private uint maxField;

        private bool manField;

        private bool ptField;

        public CT_Break()
        {
            this.idField = ((uint)(0));
            this.minField = ((uint)(0));
            this.maxField = ((uint)(0));
            this.manField = false;
            this.ptField = false;
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint min
        {
            get
            {
                return this.minField;
            }
            set
            {
                this.minField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "0")]
        public uint max
        {
            get
            {
                return this.maxField;
            }
            set
            {
                this.maxField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool man
        {
            get
            {
                return this.manField;
            }
            set
            {
                this.manField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool pt
        {
            get
            {
                return this.ptField;
            }
            set
            {
                this.ptField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PageSetup
    {

        private uint paperSizeField;

        private uint scaleField;

        private uint firstPageNumberField;

        private uint fitToWidthField;

        private uint fitToHeightField;

        private ST_PageOrder pageOrderField;

        private ST_Orientation orientationField;

        private bool usePrinterDefaultsField;

        private bool blackAndWhiteField;

        private bool draftField;

        private ST_CellComments cellCommentsField;

        private bool useFirstPageNumberField;

        private ST_PrintError errorsField;

        private uint horizontalDpiField;

        private uint verticalDpiField;

        private uint copiesField;

        private string idField;

        public CT_PageSetup()
        {
            this.paperSizeField = ((uint)(1));
            this.scaleField = ((uint)(100));
            this.firstPageNumberField = ((uint)(1));
            this.fitToWidthField = ((uint)(1));
            this.fitToHeightField = ((uint)(1));
            this.pageOrderField = ST_PageOrder.downThenOver;
            this.orientationField = ST_Orientation.@default;
            this.usePrinterDefaultsField = true;
            this.blackAndWhiteField = false;
            this.draftField = false;
            this.cellCommentsField = ST_CellComments.none;
            this.useFirstPageNumberField = false;
            this.errorsField = ST_PrintError.displayed;
            this.horizontalDpiField = ((uint)(600));
            this.verticalDpiField = ((uint)(600));
            this.copiesField = ((uint)(1));
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint paperSize
        {
            get
            {
                return this.paperSizeField;
            }
            set
            {
                this.paperSizeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "100")]
        public uint scale
        {
            get
            {
                return this.scaleField;
            }
            set
            {
                this.scaleField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint firstPageNumber
        {
            get
            {
                return this.firstPageNumberField;
            }
            set
            {
                this.firstPageNumberField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint fitToWidth
        {
            get
            {
                return this.fitToWidthField;
            }
            set
            {
                this.fitToWidthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint fitToHeight
        {
            get
            {
                return this.fitToHeightField;
            }
            set
            {
                this.fitToHeightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_PageOrder.downThenOver)]
        public ST_PageOrder pageOrder
        {
            get
            {
                return this.pageOrderField;
            }
            set
            {
                this.pageOrderField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_Orientation.@default)]
        public ST_Orientation orientation
        {
            get
            {
                return this.orientationField;
            }
            set
            {
                this.orientationField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(true)]
        public bool usePrinterDefaults
        {
            get
            {
                return this.usePrinterDefaultsField;
            }
            set
            {
                this.usePrinterDefaultsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool blackAndWhite
        {
            get
            {
                return this.blackAndWhiteField;
            }
            set
            {
                this.blackAndWhiteField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool draft
        {
            get
            {
                return this.draftField;
            }
            set
            {
                this.draftField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_CellComments.none)]
        public ST_CellComments cellComments
        {
            get
            {
                return this.cellCommentsField;
            }
            set
            {
                this.cellCommentsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(false)]
        public bool useFirstPageNumber
        {
            get
            {
                return this.useFirstPageNumberField;
            }
            set
            {
                this.useFirstPageNumberField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(ST_PrintError.displayed)]
        public ST_PrintError errors
        {
            get
            {
                return this.errorsField;
            }
            set
            {
                this.errorsField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "600")]
        public uint horizontalDpi
        {
            get
            {
                return this.horizontalDpiField;
            }
            set
            {
                this.horizontalDpiField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "600")]
        public uint verticalDpi
        {
            get
            {
                return this.verticalDpiField;
            }
            set
            {
                this.verticalDpiField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        [System.ComponentModel.DefaultValueAttribute(typeof(uint), "1")]
        public uint copies
        {
            get
            {
                return this.copiesField;
            }
            set
            {
                this.copiesField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PageOrder
    {

        /// <remarks/>
        downThenOver,

        /// <remarks/>
        overThenDown,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_Orientation
    {

        /// <remarks/>
        @default,

        /// <remarks/>
        portrait,

        /// <remarks/>
        landscape,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_CellComments
    {

        /// <remarks/>
        none,

        /// <remarks/>
        asDisplayed,

        /// <remarks/>
        atEnd,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public enum ST_PrintError
    {

        /// <remarks/>
        displayed,

        /// <remarks/>
        blank,

        /// <remarks/>
        dash,

        /// <remarks/>
        NA,
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public partial class CT_PageMargins
    {

        private double leftField;

        private double rightField;

        private double topField;

        private double bottomField;

        private double headerField;

        private double footerField;

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double left
        {
            get
            {
                return this.leftField;
            }
            set
            {
                this.leftField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double right
        {
            get
            {
                return this.rightField;
            }
            set
            {
                this.rightField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double top
        {
            get
            {
                return this.topField;
            }
            set
            {
                this.topField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double bottom
        {
            get
            {
                return this.bottomField;
            }
            set
            {
                this.bottomField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double header
        {
            get
            {
                return this.headerField;
            }
            set
            {
                this.headerField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public double footer
        {
            get
            {
                return this.footerField;
            }
            set
            {
                this.footerField = value;
            }
        }
    }
    
}
