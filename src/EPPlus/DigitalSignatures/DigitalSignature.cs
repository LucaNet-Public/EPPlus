using System;
using System.IO;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA.Signatures;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;

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

        string GetPackageObject()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<Object Id=\"idOfficeObject\">");
            sb.Append("<Manifest>");
            sb.Append("<Reference URI=\"/_rels/.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml\">");
            sb.Append("<Transforms>");
            sb.Append("<Transform Algorithm=\"http://schemas.openxmlformats.org/package/2006/RelationshipTransform\">");
            sb.Append("<mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" SourceId=\"rId1\"/>");
            sb.Append("</Transform>");
            sb.Append("<Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/>");
            sb.Append("</Transforms>");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");
            sb.Append("<DigestValue>+nAd0bim5u961Z6hkrztwiSj8HA=</DigestValue>");
            sb.Append("</Reference>");
            sb.Append("<Reference URI=\"/xl/_rels/workbook.xml.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml\">");
            sb.Append("<Transforms>");
            sb.Append("<Transform Algorithm=\"http://schemas.openxmlformats.org/package/2006/RelationshipTransform\">");
            sb.Append("<mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" SourceId=\"rId2\"/>");
            sb.Append("<mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" SourceId=\"rId1\"/>");
            sb.Append("<mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" SourceId=\"rId3\"/>");
            sb.Append("</Transform>");
            sb.Append("<Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/>");
            sb.Append("</Transforms>");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");
            sb.Append("<DigestValue>9avqBD/nBx+S4OXeMoQGXb6Ex+I=</DigestValue>");
            sb.Append("</Reference>");

            sb.Append("<Reference URI=\"/xl/styles.xml?ContentType=application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\">");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");
            sb.Append("<DigestValue>R3jSMFWoLJ87ma2wdBoixK+0JNU=</DigestValue>");
            sb.Append("</Reference>");

            sb.Append("<Reference URI=\"/xl/theme/theme1.xml?ContentType=application/vnd.openxmlformats-officedocument.theme+xml\">");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");
            sb.Append("<DigestValue>/ORquKgdTLXGvZMWJ4gf11/OKGo=</DigestValue>");
            sb.Append("</Reference>");

            sb.Append("<Reference URI=\"/xl/workbook.xml?ContentType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\">");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");
            sb.Append("<DigestValue>51TgNdwVubJZT8fUA/pVPTsIaQI=</DigestValue>");
            sb.Append("</Reference>");

            sb.Append("<Reference URI=\"/xl/worksheets/sheet1.xml?ContentType=application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\">");
            sb.Append("<DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/>");
            sb.Append("<DigestValue>kmvOJh6NwC8/DzXOVHP84iPGQys=</DigestValue>");
            sb.Append("</Reference>");

            sb.Append("</Manifest>");

            sb.Append("<SignatureProperties>");
            sb.Append("<SignatureProperty Id=\"idSignatureTime\" Target=\"#idPackageSignature\">");
            sb.Append("<mdssi:SignatureTime xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\">");
            sb.Append("<mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>");
            sb.Append("<mdssi:Value>2024-05-27T12:07:02Z</mdssi:Value>");
            sb.Append("</mdssi:SignatureTime>");
            sb.Append("</SignatureProperty>");
            sb.Append("</SignatureProperties>");

            sb.Append("</Object>");

            return sb.ToString();
        }

        string GetOfficeObject()
        {
            StringBuilder sb = new StringBuilder();

            //Placeholder data
            string commitment = "Forty-two.";
            string officeVersion = "16.0.17531/26";
            string applicationVersion = "16.0.17531";
            string numMonitors = "3";
            string hRes = "2560";
            string vRes = "1440";
            string colorDepth = "32";
            string empty = "";
            string signatureProviderDetails = "9";

            //If this is set to 2 REQUIRES additional objects idValidSigLnImg and idInvalidSigLnImg
            string signatureType = "1";

            string defaultSignatureProvider = "{00000000-0000-0000-0000-000000000000}";
            string signatureProviderId = defaultSignatureProvider;

            bool eastAsianprovider = false;
            if(eastAsianprovider)
            {
                signatureProviderId = "{000CD6A4-0000-0000-C000-000000000046}";
            }

            sb.Append("<Object Id=\"idOfficeObject\">");
            sb.Append("<SignatureProperties>");
            sb.Append("<SignatureProperty Id=\"idOfficeV1Details\" Target=\"#idPackageSignature\">");
            sb.Append("<SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\">");
            sb.Append("<SetupID></SetupID>");
            sb.Append("<SignatureText></SignatureText>");
            sb.Append("<SignatureImage />");
            sb.Append($"<SignatureComments>{commitment}</SignatureComments>");
            sb.Append($"<WindowsVersion>{Environment.OSVersion.ToString()}</WindowsVersion>");
            sb.Append($"<OfficeVersion>{officeVersion}</OfficeVersion>");
            sb.Append($"<ApplicationVersion>{applicationVersion}</ApplicationVersion>");
            sb.Append($"<Monitors>{numMonitors}</Monitors>");
            sb.Append($"<HorizontalResolution>{hRes}</HorizontalResolution>");
            sb.Append($"<VerticalResolution>{vRes}</VerticalResolution>");
            sb.Append($"<ColorDepth>{colorDepth}</ColorDepth>");
            sb.Append($"<SignatureProviderId>{signatureProviderId}</SignatureProviderId>");
            sb.Append($"<SignatureProviderUrl>{empty}</SignatureProviderUrl>");
            sb.Append($"<SignatureProviderDetails>{signatureProviderDetails}</SignatureProviderDetails>");
            sb.Append($"<SignatureType>{signatureType}</SignatureType>");
            sb.Append($"</SignatureInfoV1>");
            sb.Append("</SignatureProperty>");
            sb.Append("</SignatureProperties>");
            sb.Append("</Object>");

            return sb.ToString();
        }


        public string HashAndEncodeBytes(byte[] temp)
        {
            using (var sha1Hash = SHA1.Create())
            {
                var hash = sha1Hash.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }

        string GetKeyInfo()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<KeyInfo>");
            sb.Append("<X509Data>");
            sb.Append($"<X509Certificate>{HashAndEncodeBytes(Certificate.GetPublicKey())}</X509Certificate>");
            sb.Append("</X509Data>");
            sb.Append("</KeyInfo>");

            return sb.ToString();
        }

        internal void Save()
        {
            string parentElement = "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" Id=\"idPackageSignature\">";
            string endElement = "</Signature>";



            var keyinfo = GetKeyInfo();
            var xmlIdSignedProperties = GetQualifyingPropertiesXml();
            var officeObject = GetOfficeObject();
            var packageObject = GetPackageObject();

            var joinedString = string.Concat(parentElement, keyinfo, packageObject, officeObject, xmlIdSignedProperties, endElement);


            byte[] bytes = Encoding.UTF8.GetBytes(joinedString);

            _part.GetStream(FileMode.Create).Write(bytes, 0, bytes.Length);

            CspParameters cspParams = new()
            {
                KeyContainerName = "XML_DSIG_RSA_KEY",
            };

            RSACryptoServiceProvider rsaKey = new(cspParams);

            XmlDocument xmlDoc = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc.LoadXml(_part.Uri.AbsolutePath);

            SignedXml signedXml = new(xmlDoc)
            {
                SigningKey = rsaKey
            };

            Reference reference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idOfficeObject"
            };

            XmlDsigEnvelopedSignatureTransform env = new();
            reference.AddTransform(env);

            signedXml.AddReference(reference);

            signedXml.ComputeSignature();
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
