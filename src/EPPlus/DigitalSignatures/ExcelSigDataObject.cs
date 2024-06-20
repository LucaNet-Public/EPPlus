using System.Security.Cryptography.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ExcelSigDataObject
    {
        DataObject _obj = new DataObject();


        public ExcelSigDataObject() 
        {

        }

        public ExcelSigDataObject(DataObject obj) 
        { 
            _obj = obj;
        }

    }
}
