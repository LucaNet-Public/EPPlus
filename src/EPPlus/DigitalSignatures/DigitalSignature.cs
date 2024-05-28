using System;
using System.IO;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA.Signatures;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class DigitalSignature
    {
        ZipPackagePart _part;

        public X509Certificate2 Certificate { get; set; }
        public SignedCms Verifier { get; internal set; }
        public EPPlusSignatureContext Context { get; set; }

        internal const string PartUri = @"/_xmlsignatures/sig1.xml";

        internal DigitalSignature(ExcelWorkbook wb) 
        {
            _part = wb._package.ZipPackage.CreatePart(new Uri(PartUri, UriKind.Relative), ContentTypes.xmlSignatures);
        }

        string GetQualifyingPropertiesXml()
        {
            StringBuilder sb = new StringBuilder();

            var stringTest = Certificate.ToString(true);
            var test = Certificate.PublicKey.Key.ToXmlString(false);

            sb.Append("<Object>");
            sb.Append("<xd:QualifyingProperties xmlns:xd=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"#idPackageSignature\">");
            sb.Append("<xd:SignedProperties Id=\"idSignedProperties\">");
            sb.Append("<xd:SignedSignatureProperties>");
            sb.Append("<xd:SigningTime>");
            sb.Append($"{DateTime.Now.ToString("yyyy-MM-dd")}T{DateTime.Now.ToString("HH:mm")}Z");
            sb.Append($"</xd:SigningTime>");
            sb.Append($"<xd:SigningCertificate>");
            sb.Append("<xd:Cert>");
            sb.Append("<xd:CertDigest>");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");

            //Note that the GetCertHashString() method specifically returns SHA1 formatted hexadecimal string. https://learn.microsoft.com/en-us/previous-versions/jj611227(v=vs.113)
            sb.Append($"<DigestValue>{Convert.ToBase64String(Encoding.ASCII.GetBytes(Certificate.GetCertHashString()))}</DigestValue>");

            sb.Append("</xd:CertDigest>");
            sb.Append("<xd:IssuerSerial>");
            sb.Append($"<X509IssuerName>{Certificate.IssuerName.Name}</X509IssuerName>");
            sb.Append($"<X509SerialNumber>{Certificate.SerialNumber}</X509SerialNumber>");
            sb.Append($"</xd:IssuerSerial>");
            sb.Append("</xd:Cert>");
            sb.Append("</xd:SigningCertificate>");
            sb.Append("<xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties>");
            sb.Append("<xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId>");

            string proofOfOrigin = "http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin";
            string identifierDescription = "Created and approved this document";

            sb.Append($"<xd:Identifier>{proofOfOrigin}</xd:Identifier>");
            sb.Append($"<xd:Description>{identifierDescription}</xd:Description>");

            sb.Append("</xd:CommitmentTypeId>");
            sb.Append($"<xd:AllSignedDataObjects/>");
            sb.Append($"<xd:CommitmentTypeQualifiers>");
            string commitment = "Forty-two.";
            sb.Append($"<xd:CommitmentTypeQualifier>{commitment}</xd:CommitmentTypeQualifier>");
            sb.Append("</xd:CommitmentTypeQualifiers>");
            sb.Append("</xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties></xd:QualifyingProperties></Object>");

            return sb.ToString();
        }


        string GetKeyInfo()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<KeyInfo>");
            sb.Append("<X509Data>");
            sb.Append($"<X509Certificate>{Certificate.GetPublicKeyString()}</X509Certificate>");
            sb.Append("</X509Data>");
            sb.Append("</KeyInfo>");

            return sb.ToString();
        }

        internal void Save()
        {
            var keyinfo = GetKeyInfo();
            var xmlIdSignedProperties = GetQualifyingPropertiesXml();

            var joinedString = string.Concat(keyinfo, xmlIdSignedProperties);

            byte[] bytes = Encoding.UTF8.GetBytes(joinedString);

            _part.GetStream(FileMode.Create).Write(bytes, 0, bytes.Length);
        }

        internal void CreateSignature()
        {
            byte[] certStore = CertUtil.GetSerializedCertStore(Certificate.RawData);
            if (Certificate == null)
            {
                SignaturePartUtil.DeleteParts(_part);
                return;
            }

            if (Certificate.HasPrivateKey == false)    //No signature. Remove any Signature part
            {
                var storeCert = CertUtil.GetCertificate(Certificate.Thumbprint);
                if (storeCert != null)
                {
                    Certificate = storeCert;
                }
                else
                {
                    SignaturePartUtil.DeleteParts(_part);
                    return;
                }
            }

            using (var ms = RecyclableMemory.GetStream())
            {
                var bw = new BinaryWriter(ms);
                //Verifier = CertUtil.SignProject(project, this, Context);
                var cert = Verifier.Encode();
                var signatureBytes = CertUtil.CreateBinarySignature(ms, bw, certStore, cert);
                //_part = SignaturePartUtil.GetPart(project, this);
                _part.GetStream(FileMode.Create).Write(signatureBytes, 0, signatureBytes.Length);
            }
        }
    }
}
