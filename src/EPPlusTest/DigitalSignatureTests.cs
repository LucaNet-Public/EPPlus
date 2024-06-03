using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.Cryptography;
using System.Xml;
using System.Security.Cryptography.Xml;
using System.IO;
using System;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml.Drawing;

//REMEMBER:
//1. Cannonize
//2. Hash data

namespace EPPlusTest
{
    [TestClass]
    public class DigitalSignatureTests : TestBase
    {
        public string HashAndEncodeBytes(byte[] temp)
        {
            using (var sha1Hash = SHA1.Create())
            {
                var hash = sha1Hash.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }

        [TestMethod]
        public void VerifyTheory2()
        {
            //CspParameters cspParams = new()
            //{
            //    KeyContainerName = "XML_DSIG_RSA_KEY",
            //};

            // RSACryptoServiceProvider rsaKey = new(cspParams);

            RSACryptoServiceProvider rsaKey = new();

            XmlDocument xmlDoc = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc.Load("C:\\epplusTest\\Workbooks\\idOfficeSeparateNew.xml");

            SignedXml signedXml = new(xmlDoc)
            {
                SigningKey = rsaKey
            };

            Reference reference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idOfficeObject"
            };

            reference.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            var idElement = signedXml.GetIdElement(xmlDoc, "idOfficeObject");

            signedXml.Signature.Id = "idPackageSignature";

            //Verify file from earlier ------------
            XmlDocument earlyXmlDoc = new()
            {
                PreserveWhitespace = true,
            };
            earlyXmlDoc.Load("C:\\epplusTest\\Workbooks\\sig1TestFile.xml");
            SignedXml signedXmlEarly = new(earlyXmlDoc)
            {
                SigningKey = rsaKey
            };

            var earlyIdElement = signedXmlEarly.GetIdElement(earlyXmlDoc, "idOfficeObject");

            Assert.AreEqual(earlyIdElement.OuterXml, idElement.OuterXml);
            //------End of verify-----------

            signedXml.AddReference(reference);

            signedXml.ComputeSignature();

            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            XmlElement xmlDigitalSignature = signedXml.GetXml();

            XmlDocument doc = new XmlDocument()
            {
                PreserveWhitespace = true,
            };

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            doc.DocumentElement?.AppendChild(doc.ImportNode(xmlDigitalSignature, true));

            doc.Save("C:\\epplusTest\\Workbooks\\newVersion.xml");
        }

        [TestMethod]
        public void SignAsExcelDoes()
        {
            //CspParameters cspParams = new()
            //{
            //    KeyContainerName = "XML_DSIG_RSA_KEY",
            //};

           // RSACryptoServiceProvider rsaKey = new(cspParams);


            RSACryptoServiceProvider rsaKey = new();

            XmlDocument xmlDoc = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc.Load("C:\\epplusTest\\Workbooks\\sig1TestFile.xml");

            SignedXml signedXml = new(xmlDoc)
            {
                SigningKey = rsaKey
            };

            Reference reference = new()
            {
                Type = "http://www.w3.org/2000/09/xmldsig#Object",
                Uri = "#idOfficeObject"
            };

            reference.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            var idElement = signedXml.GetIdElement(xmlDoc, "idOfficeObject");

            XmlDocument xmlDoc3 = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc3.Load("C:\\epplusTest\\Workbooks\\newXml.xml");


            signedXml.Signature.Id = "idPackageSignature";

            XmlDocument doc = new XmlDocument()
            {
                PreserveWhitespace = true,
            };

            XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmlDeclaration, root);

            //Breaking out idOffice part of file to new file for later verification.
            var aString = idElement.OuterXml;
            File.WriteAllText("C:\\epplusTest\\Workbooks\\idOfficeSeparateNew.xml", aString, System.Text.Encoding.UTF8);

            signedXml.AddReference(reference);

            signedXml.ComputeSignature();

            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            XmlElement xmlDigitalSignature = signedXml.GetXml();

            XmlDocument xmlDoc2 = new()
            {
                PreserveWhitespace = true,
            };
            xmlDoc2.Load("C:\\epplusTest\\Workbooks\\newXml.xml");

            xmlDoc2.DocumentElement?.AppendChild(xmlDoc2.ImportNode(xmlDigitalSignature, true));

            var listNodes = xmlDigitalSignature.GetElementsByTagName("DigestValue");
            var node1 = listNodes[0];

            Assert.AreEqual("Dwx/mtIT+lffP980qEOPVRJX41k=", node1.InnerText);
            //var stringTest = System.Text.Encoding.UTF8.GetString(reference.DigestValue);
            //Assert.AreEqual("Dwx/mtIT+lffP980qEOPVRJX41k=", stringTest);

            xmlDoc2.Save("C:\\epplusTest\\Workbooks\\newVersionExcelBased.xml");
        }

        [TestMethod]
        public void SHA1Test()
        {
            var bytes = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\sheet1.xml");
            string hash = HashAndEncodeBytes(bytes);

            Assert.AreEqual("5dK/Tn8G0h7N8XnQ6PO7YcqoOWY=", hash);
        }

        [TestMethod]
        public void VerifyTheory()
        {
            string pObject = "<Object Id=\"idOfficeObject\" xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignatureProperties><SignatureProperty Id=\"idOfficeV1Details\" Target=\"#idPackageSignature\"><SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\"><SetupID></SetupID><SignatureText></SignatureText><SignatureImage /><SignatureComments>Forty-two.</SignatureComments><WindowsVersion>10.0</WindowsVersion><OfficeVersion>16.0.17531/26</OfficeVersion><ApplicationVersion>16.0.17531</ApplicationVersion><Monitors>3</Monitors><HorizontalResolution>2560</HorizontalResolution><VerticalResolution>1440</VerticalResolution><ColorDepth>32</ColorDepth><SignatureProviderId>{00000000-0000-0000-0000-000000000000}</SignatureProviderId><SignatureProviderUrl></SignatureProviderUrl><SignatureProviderDetails>9</SignatureProviderDetails><SignatureType>1</SignatureType></SignatureInfoV1></SignatureProperty></SignatureProperties></Object>";
            File.WriteAllText("C:\\epplusTest\\Workbooks\\pObjectTest.xml", pObject, System.Text.Encoding.UTF8);
            var test = HashAndEncodeBytes(System.Text.Encoding.UTF8.GetBytes(pObject));

            var readData = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\idOfficeClean.xml");
            var res = HashAndEncodeBytes(readData);

            var officeObj = "<Object Id=\"idOfficeObject\"><SignatureProperties><SignatureProperty Id=\"idOfficeV1Details\" Target=\"#idPackageSignature\"><SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\"><SetupID></SetupID><SignatureText></SignatureText><SignatureImage/><SignatureComments>Forty-two.</SignatureComments><WindowsVersion>10.0</WindowsVersion><OfficeVersion>16.0.17531/26</OfficeVersion><ApplicationVersion>16.0.17531</ApplicationVersion><Monitors>3</Monitors><HorizontalResolution>2560</HorizontalResolution><VerticalResolution>1440</VerticalResolution><ColorDepth>32</ColorDepth><SignatureProviderId>{00000000-0000-0000-0000-000000000000}</SignatureProviderId><SignatureProviderUrl></SignatureProviderUrl><SignatureProviderDetails>9</SignatureProviderDetails><SignatureType>1</SignatureType></SignatureInfoV1></SignatureProperty></SignatureProperties></Object><Object><xd:QualifyingProperties xmlns:xd=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"#idPackageSignature\"><xd:SignedProperties Id=\"idSignedProperties\"><xd:SignedSignatureProperties><xd:SigningTime>2024-05-27T12:07:02Z</xd:SigningTime><xd:SigningCertificate><xd:Cert><xd:CertDigest><DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha1\"/><DigestValue>w9iTMIvTXcdRc9G38Pp1Njb/HPE=</DigestValue></xd:CertDigest><xd:IssuerSerial><X509IssuerName>CN=OssianEdström</X509IssuerName><X509SerialNumber>38225183535545048482234589307877617536</X509SerialNumber></xd:IssuerSerial></xd:Cert></xd:SigningCertificate><xd:SignaturePolicyIdentifier><xd:SignaturePolicyImplied/></xd:SignaturePolicyIdentifier></xd:SignedSignatureProperties><xd:SignedDataObjectProperties><xd:CommitmentTypeIndication><xd:CommitmentTypeId><xd:Identifier>http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin</xd:Identifier><xd:Description>Created and approved this document</xd:Description></xd:CommitmentTypeId><xd:AllSignedDataObjects/><xd:CommitmentTypeQualifiers><xd:CommitmentTypeQualifier>Forty-two.</xd:CommitmentTypeQualifier></xd:CommitmentTypeQualifiers></xd:CommitmentTypeIndication></xd:SignedDataObjectProperties></xd:SignedProperties></xd:QualifyingProperties></Object>";
            byte[] byteTest = System.Text.Encoding.Default.GetBytes(officeObj);

            var readDataOffice = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\packageObject.xml");
            var officeObject = File.ReadAllBytes("C:\\epplusTest\\Workbooks\\idOfficeObject.xml");

            var testDataOffice = HashAndEncodeBytes(byteTest);
            var testnytt = HashAndEncodeBytes(readDataOffice);
            var officeObjectStuff = HashAndEncodeBytes(officeObject);

            Assert.AreEqual("Dwx/mtIT+lffP980qEOPVRJX41k=", officeObjectStuff);

            Assert.AreEqual("kxA0qm2FwPZvNmtI22ItXRQHlVs=", res);

            byte[] data = Convert.FromBase64String("kxA0qm2FwPZvNmtI22ItXRQHlVs=");

            string decodedString = System.Text.Encoding.UTF8.GetString(data);
        }


        [TestMethod]
        public void DigitallySignDoc()
        {
            using (var p = OpenTemplatePackage("UnsignedWBEmpty.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
                    {
                        p.Workbook.DigSignature.Certificate = cert;
                        break;
                    }
                }

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void VBASignTest()
        {
            using (var p = OpenPackage("VbaTest.xlsm", true))
            {
                ExcelPackage pck = new ExcelPackage();
                //Add a worksheet.
                var ws = pck.Workbook.Worksheets.Add("VBA Sample");
                ws.Drawings.AddShape("VBASampleRect", eShapeStyle.RoundRect);
                //Create a vba project             
                pck.Workbook.CreateVBAProject();
                //Now add some code to update the text of the shape...
                var sb = new StringBuilder();
                sb.AppendLine("Private Sub Workbook_Open()");
                sb.AppendLine("    [VBA Sample].Shapes(\"VBASampleRect\").TextEffect.Text = \"This text is set from VBA!\"");
                sb.AppendLine("End Sub");
                pck.Workbook.CodeModule.Code = sb.ToString();

                X509Store store = new X509Store(StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                foreach (var cert in store.Certificates)
                {
                    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
                    {
                        pck.Workbook.VbaProject.Signature.Certificate = cert;
                        break;
                    }
                }

                //And Save as xlsm
                pck.SaveAs(new FileInfo(@"C:\epplusTest\Testoutput" + @"\VbaTest.xlsm"));
            }
        }
    }
}
