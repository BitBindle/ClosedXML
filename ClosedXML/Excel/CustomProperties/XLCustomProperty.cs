using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLCustomProperty : IXLCustomProperty
    {
        private readonly XLWorkbook _workbook;

        private String name;

        private Object _value;

        public XLCustomProperty(XLWorkbook workbook, String name, Object value)
        {
            _workbook = workbook;
            this.name = string.Empty;
            Name = name;
            _value = value ?? throw new ArgumentNullException(nameof(value));
        }

        #region IXLCustomProperty Members

        public String Name
        {
            get { return name; }
            set
            {
                if (name == value) return;

                if (_workbook.CustomProperties.Any(t => t.Name == value))
                    throw new ArgumentException(
                        String.Format("This workbook already contains a custom property named '{0}'", value));

                name = value;
            }
        }

        public XLCustomPropertyType Type
        {
            get
            {
                if (Value is DateTime)
                    return XLCustomPropertyType.Date;

                if (Value is Boolean)
                    return XLCustomPropertyType.Boolean;

                if (Double.TryParse(Value.ToString(), out Double dTest))
                    return XLCustomPropertyType.Number;

                return XLCustomPropertyType.Text;
            }
        }

        public Object Value
        {
            get => _value;
            set => _value = value ?? throw new ArgumentNullException(nameof(value));
        }

        public T GetValue<T>()
        {
            return (T)Convert.ChangeType(Value, typeof(T));
        }

        #endregion
    }
}
