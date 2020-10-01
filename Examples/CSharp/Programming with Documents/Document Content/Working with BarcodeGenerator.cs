﻿using System;
using System.Drawing;
using Aspose.BarCode.Generation;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_with_Documents.Document_Content
{
    class WorkingWithBarcodeGenerator : TestDataHelper
    {
        [Test]
        public static void GenerateACustomBarCodeImage()
        {
            //ExStart:GenerateACustomBarCodeImage
            Document doc = new Document(MyDir + "Field sample - BARCODE.docx");

            // Set custom barcode generator
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            doc.Save(ArtifactsDir + "GenerateACustomBarCodeImage.pdf");
            //ExEnd:GenerateACustomBarCodeImage
        }
    }

    //ExStart:GenerateACustomBarCodeImage_IBarcodeGenerator
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        /// <summary>
        /// Converts barcode image height from Word units to Aspose.BarCode units.
        /// </summary>
        /// <param name="heightInTwipsString"></param>
        /// <returns></returns>
        private static float ConvertSymbolHeight(string heightInTwipsString)
        {
            // Input value is in 1/1440 inches (twips)
            int.TryParse(heightInTwipsString, out int heightInTwips);

            if (heightInTwips == int.MinValue)
                throw new Exception("Error! Incorrect height - " + heightInTwipsString + ".");

            // Convert to mm
            return (float) (heightInTwips * 25.4 / 1440);
        }

        /// <summary>
        /// Converts barcode image color from Word to Aspose.BarCode.
        /// </summary>
        /// <param name="inputColor"></param>
        /// <returns></returns>
        private static Color ConvertColor(string inputColor)
        {
            // Input should be from "0x000000" to "0xFFFFFF"
            int.TryParse(inputColor.Replace("0x", ""), out int color);

            if (color == int.MinValue)
                throw new Exception("Error! Incorrect color - " + inputColor + ".");

            return Color.FromArgb(color >> 16, (color & 0xFF00) >> 8, color & 0xFF);

            // Backword conversion -
            //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
        }

        /// <summary>
        /// Converts bar code scaling factor from percents to float.
        /// </summary>
        /// <param name="scalingFactor"></param>
        /// <returns></returns>
        private static float ConvertScalingFactor(string scalingFactor)
        {
            bool isParsed = false;
            int.TryParse(scalingFactor, out int percents);

            if (percents != int.MinValue)
            {
                if (percents >= 10 && percents <= 10000)
                    isParsed = true;
            }

            if (!isParsed)
                throw new Exception("Error! Incorrect scaling factor - " + scalingFactor + ".");

            return percents / 100.0f;
        }

        /// <summary>
        /// Implementation of the GetBarCodeImage() method for IBarCodeGenerator interface.
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public Image GetBarcodeImage(Words.Fields.BarcodeParameters parameters)
        {
            if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
                return null;

            string type = parameters.BarcodeType.ToUpper();
            SymbologyEncodeType encodeType = EncodeTypes.None;

            switch (type)
            {
                case "QR":
                    encodeType = EncodeTypes.QR;
                    break;
                case "CODE128":
                    encodeType = EncodeTypes.Code128;
                    break;
                case "CODE39":
                    encodeType = EncodeTypes.Code39Standard;
                    break;
                case "EAN8":
                    encodeType = EncodeTypes.EAN8;
                    break;
                case "EAN13":
                    encodeType = EncodeTypes.EAN13;
                    break;
                case "UPCA":
                    encodeType = EncodeTypes.UPCA;
                    break;
                case "UPCE":
                    encodeType = EncodeTypes.UPCE;
                    break;
                case "ITF14":
                    encodeType = EncodeTypes.ITF14;
                    break;
                case "CASE":
                    encodeType = EncodeTypes.None;
                    break;
            }

            if (encodeType.Equals(EncodeTypes.None))
                return null;

            BarcodeGenerator generator = new BarcodeGenerator(encodeType);
            generator.CodeText = parameters.BarcodeValue;

            if (encodeType.Equals(EncodeTypes.QR))
                generator.Parameters.Barcode.CodeTextParameters.TwoDDisplayText = parameters.BarcodeValue;

            if (parameters.ForegroundColor != null)
                generator.Parameters.Barcode.BarColor = ConvertColor(parameters.ForegroundColor);

            if (parameters.BackgroundColor != null)
                generator.Parameters.BackColor = ConvertColor(parameters.BackgroundColor);

            if (parameters.SymbolHeight != null)
            {
                generator.Parameters.ImageHeight.Millimeters = ConvertSymbolHeight(parameters.SymbolHeight);
                generator.Parameters.AutoSizeMode = AutoSizeMode.Nearest;
            }

            generator.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;

            if (parameters.DisplayText)
                generator.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

            generator.Parameters.CaptionAbove.Text = "";

            const float scale = 1.0f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
            float xdim = 1.0f;

            if (encodeType.Equals(EncodeTypes.QR))
            {
                generator.Parameters.Barcode.AutoSizeMode = AutoSizeMode.Nearest;
                generator.Parameters.Barcode.BarCodeWidth.Millimeters *= scale;
                generator.Parameters.Barcode.BarCodeHeight.Millimeters =
                    generator.Parameters.Barcode.BarCodeWidth.Millimeters;
                xdim = generator.Parameters.Barcode.BarCodeHeight.Millimeters / 25;
                generator.Parameters.Barcode.XDimension.Millimeters =
                    generator.Parameters.Barcode.BarHeight.Millimeters = xdim;
            }

            if (parameters.ScalingFactor != null)
            {
                float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
                generator.Parameters.Barcode.BarCodeHeight.Millimeters *= scalingFactor;
                if (encodeType.Equals(EncodeTypes.QR))
                {
                    generator.Parameters.Barcode.BarCodeWidth.Millimeters =
                        generator.Parameters.Barcode.BarCodeHeight.Millimeters;
                    generator.Parameters.Barcode.XDimension.Millimeters =
                        generator.Parameters.Barcode.BarHeight.Millimeters = xdim * scalingFactor;
                }

                generator.Parameters.AutoSizeMode = AutoSizeMode.Nearest;
            }

            return generator.GenerateBarCodeImage();
        }

        public Image GetOldBarcodeImage(Words.Fields.BarcodeParameters parameters)
        {
            throw new NotImplementedException();
        }
    }
    //ExEnd:GenerateACustomBarCodeImage_IBarcodeGenerator
}