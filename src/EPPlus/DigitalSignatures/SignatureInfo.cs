using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    public class SignatureInfo
    {
        internal SignatureInfo(bool eastAsianProvider = false) 
        {
            string defaultSignatureProvider = "{00000000-0000-0000-0000-000000000000}";
            SignatureProviderID = defaultSignatureProvider;
            WindowsVersion = Environment.OSVersion.ToString();

            if (eastAsianProvider)
            {
                SignatureProviderID = "{000CD6A4-0000-0000-C000-000000000046}";
            }
        }

        //Required children
        public string SetUpId;
        public string SignatureText;
        //Base64 binary
        public string SignatureImage;
        public string SignatureComments;
        public string WindowsVersion;
        public string OfficeVersion;
        public string ApplicationVersion;
        public uint Monitors;
        public uint HorizontalResolution;
        public uint VerticalResolution;
        public uint ColorDepth;
        public string SignatureProviderID;
        public string SignatureProviderUrl;
        public int SignatureProviderDetails;
        public DigitalSignatureType SignatureType;
        //Optional children
        public string DeletegateSuggesterSigner;
        public string DelegateSuggesterSigner2;
        public string DelegateSuggestedSignerEmail;
        public Uri ManifestHashAlgorithm;
        //Optional children2
        public string Address1;
        public string Address2;

        internal string GetXml()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\">");
            sb.Append($"<SetupID>{SetUpId}</SetupID>");
            sb.Append($"<SignatureText>{SignatureText}</SignatureText>");
            sb.Append($"<SignatureImage>{SignatureImage}</SignatureImage>");
            sb.Append($"<SignatureComments>{SignatureComments}</SignatureComments>");
            sb.Append($"<WindowsVersion>{WindowsVersion}</WindowsVersion>");
            sb.Append($"<OfficeVersion>{OfficeVersion}</OfficeVersion>");
            sb.Append($"<ApplicationVersion>{ApplicationVersion}</ApplicationVersion>");
            sb.Append($"<Monitors>{Monitors}</Monitors>");
            sb.Append($"<HorizontalResolution>{HorizontalResolution}</HorizontalResolution>");
            sb.Append($"<VerticalResolution>{VerticalResolution}</VerticalResolution>");
            sb.Append($"<ColorDepth>{ColorDepth}</ColorDepth>");
            sb.Append($"<SignatureProviderId>{SignatureProviderID}</SignatureProviderId>");
            sb.Append($"<SignatureProviderUrl>{SignatureProviderUrl}</SignatureProviderUrl>");
            sb.Append($"<SignatureProviderDetails>{SignatureProviderDetails}</SignatureProviderDetails>");
            sb.Append($"<SignatureType>{SignatureType}</SignatureType>");
            sb.Append($"</SignatureInfoV1>");

            return sb.ToString();
        }
    }
}
