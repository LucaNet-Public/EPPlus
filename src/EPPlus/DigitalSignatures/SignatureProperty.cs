using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    public class SignatureProperty
    {
        internal string Target;
        internal string Id;

        internal SignatureProperty(string target, string id) 
        {
            Target = target;
            Id = id;
            SignatureInfoV1 = new SignatureInfo();
        }

        public SignatureInfo SignatureInfoV1 = null;
        public SignatureInfo SignatureInfoV2 = null;

    }
}
