using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using OptimaJet.DWKit.Core;
using OptimaJet.DWKit.Core.Model;
using static OptimaJet.DWKit.Application.Utilities.Tools;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using Order = DocumentFormat.OpenXml.Drawing.Charts.Order;
public class sampleChart
{
    public ChartSpace chart1()
    {
        ChartSpace chartSpace1 = new ChartSpace();
        chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
        Date1904 date19041 = new Date1904() { Val = false };
        EditingLanguage editingLanguage1 = new EditingLanguage() { Val = "en-US" };
        RoundedCorners roundedCorners1 = new RoundedCorners() { Val = false };

        AlternateContent alternateContent1 = new AlternateContent();
        alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

        AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
        alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
        C14.Style style1 = new C14.Style() { Val = 102 };

        alternateContentChoice1.Append(style1);

        AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
        Style style2 = new Style() { Val = 2 };

        alternateContentFallback1.Append(style2);

        alternateContent1.Append(alternateContentChoice1);
        alternateContent1.Append(alternateContentFallback1);

        Chart chart1 = new Chart();

        Title title1 = new Title();
        Overlay overlay1 = new Overlay() { Val = false };

        ChartShapeProperties chartShapeProperties1 = new ChartShapeProperties();
        A.NoFill noFill1 = new A.NoFill();

        A.Outline outline1 = new A.Outline();
        A.NoFill noFill2 = new A.NoFill();

        outline1.Append(noFill2);
        A.EffectList effectList1 = new A.EffectList();

        chartShapeProperties1.Append(noFill1);
        chartShapeProperties1.Append(outline1);
        chartShapeProperties1.Append(effectList1);

        TextProperties textProperties1 = new TextProperties();
        A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
        A.ListStyle listStyle1 = new A.ListStyle();

        A.Paragraph paragraph1 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };

        A.SolidFill solidFill1 = new A.SolidFill();

        A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor1.Append(luminanceModulation1);
        schemeColor1.Append(luminanceOffset1);

        solidFill1.Append(schemeColor1);
        A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties1.Append(solidFill1);
        defaultRunProperties1.Append(latinFont1);
        defaultRunProperties1.Append(eastAsianFont1);
        defaultRunProperties1.Append(complexScriptFont1);

        paragraphProperties1.Append(defaultRunProperties1);
        A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph1.Append(paragraphProperties1);
        paragraph1.Append(endParagraphRunProperties1);

        textProperties1.Append(bodyProperties1);
        textProperties1.Append(listStyle1);
        textProperties1.Append(paragraph1);

        title1.Append(overlay1);
        title1.Append(chartShapeProperties1);
        title1.Append(textProperties1);
        AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };

        PlotArea plotArea1 = new PlotArea();
        Layout layout1 = new Layout();

        BarChart barChart1 = new BarChart();
        BarDirection barDirection1 = new BarDirection() { Val = BarDirectionValues.Column };
        BarGrouping barGrouping1 = new BarGrouping() { Val = BarGroupingValues.Stacked };
        VaryColors varyColors1 = new VaryColors() { Val = false };

        BarChartSeries barChartSeries1 = new BarChartSeries();
        Index index1 = new Index() { Val = (UInt32Value)0U };
        Order order1 = new Order() { Val = (UInt32Value)0U };

        SeriesText seriesText1 = new SeriesText();

        StringReference stringReference1 = new StringReference();
        Formula formula1 = new Formula();
        formula1.Text = "Sheet1!$B$1";

        StringCache stringCache1 = new StringCache();
        PointCount pointCount1 = new PointCount() { Val = (UInt32Value)1U };

        StringPoint stringPoint1 = new StringPoint() { Index = (UInt32Value)0U };
        NumericValue numericValue1 = new NumericValue();
        numericValue1.Text = "Claimed to Date";

        stringPoint1.Append(numericValue1);

        stringCache1.Append(pointCount1);
        stringCache1.Append(stringPoint1);

        stringReference1.Append(formula1);
        stringReference1.Append(stringCache1);

        seriesText1.Append(stringReference1);

        ChartShapeProperties chartShapeProperties2 = new ChartShapeProperties();

        A.SolidFill solidFill2 = new A.SolidFill();
        A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

        solidFill2.Append(schemeColor2);

        A.Outline outline2 = new A.Outline();
        A.NoFill noFill3 = new A.NoFill();

        outline2.Append(noFill3);
        A.EffectList effectList2 = new A.EffectList();

        chartShapeProperties2.Append(solidFill2);
        chartShapeProperties2.Append(outline2);
        chartShapeProperties2.Append(effectList2);
        InvertIfNegative invertIfNegative1 = new InvertIfNegative() { Val = false };

        DataLabels dataLabels1 = new DataLabels();

        ChartShapeProperties chartShapeProperties3 = new ChartShapeProperties();
        A.NoFill noFill4 = new A.NoFill();

        A.Outline outline3 = new A.Outline();
        A.NoFill noFill5 = new A.NoFill();

        outline3.Append(noFill5);
        A.EffectList effectList3 = new A.EffectList();

        chartShapeProperties3.Append(noFill4);
        chartShapeProperties3.Append(outline3);
        chartShapeProperties3.Append(effectList3);

        TextProperties textProperties2 = new TextProperties();

        A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
        A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

        bodyProperties2.Append(shapeAutoFit1);
        A.ListStyle listStyle2 = new A.ListStyle();

        A.Paragraph paragraph2 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

        A.SolidFill solidFill3 = new A.SolidFill();

        A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 75000 };
        A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 25000 };

        schemeColor3.Append(luminanceModulation2);
        schemeColor3.Append(luminanceOffset2);

        solidFill3.Append(schemeColor3);
        A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties2.Append(solidFill3);
        defaultRunProperties2.Append(latinFont2);
        defaultRunProperties2.Append(eastAsianFont2);
        defaultRunProperties2.Append(complexScriptFont2);

        paragraphProperties2.Append(defaultRunProperties2);
        A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph2.Append(paragraphProperties2);
        paragraph2.Append(endParagraphRunProperties2);

        textProperties2.Append(bodyProperties2);
        textProperties2.Append(listStyle2);
        textProperties2.Append(paragraph2);
        ShowLegendKey showLegendKey1 = new ShowLegendKey() { Val = false };
        ShowValue showValue1 = new ShowValue() { Val = true };
        ShowCategoryName showCategoryName1 = new ShowCategoryName() { Val = false };
        ShowSeriesName showSeriesName1 = new ShowSeriesName() { Val = false };
        ShowPercent showPercent1 = new ShowPercent() { Val = false };
        ShowBubbleSize showBubbleSize1 = new ShowBubbleSize() { Val = false };
        ShowLeaderLines showLeaderLines1 = new ShowLeaderLines() { Val = false };

        DLblsExtensionList dLblsExtensionList1 = new DLblsExtensionList();

        DLblsExtension dLblsExtension1 = new DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
        dLblsExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
        C15.ShowLeaderLines showLeaderLines2 = new C15.ShowLeaderLines() { Val = true };

        C15.LeaderLines leaderLines1 = new C15.LeaderLines();

        ChartShapeProperties chartShapeProperties4 = new ChartShapeProperties();

        A.Outline outline4 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill4 = new A.SolidFill();

        A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 35000 };
        A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 65000 };

        schemeColor4.Append(luminanceModulation3);
        schemeColor4.Append(luminanceOffset3);

        solidFill4.Append(schemeColor4);
        A.Round round1 = new A.Round();

        outline4.Append(solidFill4);
        outline4.Append(round1);
        A.EffectList effectList4 = new A.EffectList();

        chartShapeProperties4.Append(outline4);
        chartShapeProperties4.Append(effectList4);

        leaderLines1.Append(chartShapeProperties4);

        dLblsExtension1.Append(showLeaderLines2);
        dLblsExtension1.Append(leaderLines1);

        dLblsExtensionList1.Append(dLblsExtension1);

        dataLabels1.Append(chartShapeProperties3);
        dataLabels1.Append(textProperties2);
        dataLabels1.Append(showLegendKey1);
        dataLabels1.Append(showValue1);
        dataLabels1.Append(showCategoryName1);
        dataLabels1.Append(showSeriesName1);
        dataLabels1.Append(showPercent1);
        dataLabels1.Append(showBubbleSize1);
        dataLabels1.Append(showLeaderLines1);
        dataLabels1.Append(dLblsExtensionList1);

        CategoryAxisData categoryAxisData1 = new CategoryAxisData();

        StringReference stringReference2 = new StringReference();
        Formula formula2 = new Formula();
        formula2.Text = "Sheet1!$A$2:$A$7";

        StringCache stringCache2 = new StringCache();
        PointCount pointCount2 = new PointCount() { Val = (UInt32Value)6U };

        StringPoint stringPoint2 = new StringPoint() { Index = (UInt32Value)0U };
        NumericValue numericValue2 = new NumericValue();
        numericValue2.Text = "Schedule 1.0";

        stringPoint2.Append(numericValue2);

        StringPoint stringPoint3 = new StringPoint() { Index = (UInt32Value)1U };
        NumericValue numericValue3 = new NumericValue();
        numericValue3.Text = "Schedule 2.0";

        stringPoint3.Append(numericValue3);

        StringPoint stringPoint4 = new StringPoint() { Index = (UInt32Value)2U };
        NumericValue numericValue4 = new NumericValue();
        numericValue4.Text = "Schedule 3.0";

        stringPoint4.Append(numericValue4);

        StringPoint stringPoint5 = new StringPoint() { Index = (UInt32Value)3U };
        NumericValue numericValue5 = new NumericValue();
        numericValue5.Text = "Schedule 4.0";

        stringPoint5.Append(numericValue5);

        StringPoint stringPoint6 = new StringPoint() { Index = (UInt32Value)4U };
        NumericValue numericValue6 = new NumericValue();
        numericValue6.Text = "Schedule 5.0";

        stringPoint6.Append(numericValue6);

        StringPoint stringPoint7 = new StringPoint() { Index = (UInt32Value)5U };
        NumericValue numericValue7 = new NumericValue();
        numericValue7.Text = "Schedule 6.0";

        stringPoint7.Append(numericValue7);

        stringCache2.Append(pointCount2);
        stringCache2.Append(stringPoint2);
        stringCache2.Append(stringPoint3);
        stringCache2.Append(stringPoint4);
        stringCache2.Append(stringPoint5);
        stringCache2.Append(stringPoint6);
        stringCache2.Append(stringPoint7);

        stringReference2.Append(formula2);
        stringReference2.Append(stringCache2);

        categoryAxisData1.Append(stringReference2);

        Values values1 = new Values();

        NumberReference numberReference1 = new NumberReference();
        Formula formula3 = new Formula();
        formula3.Text = "Sheet1!$B$2:$B$7";

        NumberingCache numberingCache1 = new NumberingCache();
        FormatCode formatCode1 = new FormatCode();
        formatCode1.Text = "General";
        PointCount pointCount3 = new PointCount() { Val = (UInt32Value)6U };

        NumericPoint numericPoint1 = new NumericPoint() { Index = (UInt32Value)0U };
        NumericValue numericValue8 = new NumericValue();
        numericValue8.Text = "4.3";

        numericPoint1.Append(numericValue8);

        NumericPoint numericPoint2 = new NumericPoint() { Index = (UInt32Value)1U };
        NumericValue numericValue9 = new NumericValue();
        numericValue9.Text = "2.5";

        numericPoint2.Append(numericValue9);

        NumericPoint numericPoint3 = new NumericPoint() { Index = (UInt32Value)2U };
        NumericValue numericValue10 = new NumericValue();
        numericValue10.Text = "3.5";

        numericPoint3.Append(numericValue10);

        NumericPoint numericPoint4 = new NumericPoint() { Index = (UInt32Value)3U };
        NumericValue numericValue11 = new NumericValue();
        numericValue11.Text = "4.5";

        numericPoint4.Append(numericValue11);

        NumericPoint numericPoint5 = new NumericPoint() { Index = (UInt32Value)4U };
        NumericValue numericValue12 = new NumericValue();
        numericValue12.Text = "4.5";

        numericPoint5.Append(numericValue12);

        NumericPoint numericPoint6 = new NumericPoint() { Index = (UInt32Value)5U };
        NumericValue numericValue13 = new NumericValue();
        numericValue13.Text = "4.5";

        numericPoint6.Append(numericValue13);

        numberingCache1.Append(formatCode1);
        numberingCache1.Append(pointCount3);
        numberingCache1.Append(numericPoint1);
        numberingCache1.Append(numericPoint2);
        numberingCache1.Append(numericPoint3);
        numberingCache1.Append(numericPoint4);
        numberingCache1.Append(numericPoint5);
        numberingCache1.Append(numericPoint6);

        numberReference1.Append(formula3);
        numberReference1.Append(numberingCache1);

        values1.Append(numberReference1);

        BarSerExtensionList barSerExtensionList1 = new BarSerExtensionList();

        BarSerExtension barSerExtension1 = new BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-EC2A-431D-B866-E1E729B46B4D}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

        barSerExtension1.Append(openXmlUnknownElement1);

        barSerExtensionList1.Append(barSerExtension1);

        barChartSeries1.Append(index1);
        barChartSeries1.Append(order1);
        barChartSeries1.Append(seriesText1);
        barChartSeries1.Append(chartShapeProperties2);
        barChartSeries1.Append(invertIfNegative1);
        barChartSeries1.Append(dataLabels1);
        barChartSeries1.Append(categoryAxisData1);
        barChartSeries1.Append(values1);
        barChartSeries1.Append(barSerExtensionList1);

        BarChartSeries barChartSeries2 = new BarChartSeries();
        Index index2 = new Index() { Val = (UInt32Value)1U };
        Order order2 = new Order() { Val = (UInt32Value)1U };

        SeriesText seriesText2 = new SeriesText();

        StringReference stringReference3 = new StringReference();
        Formula formula4 = new Formula();
        formula4.Text = "Sheet1!$C$1";

        StringCache stringCache3 = new StringCache();
        PointCount pointCount4 = new PointCount() { Val = (UInt32Value)1U };

        StringPoint stringPoint8 = new StringPoint() { Index = (UInt32Value)0U };
        NumericValue numericValue14 = new NumericValue();
        numericValue14.Text = "Contract Work to Claim";

        stringPoint8.Append(numericValue14);

        stringCache3.Append(pointCount4);
        stringCache3.Append(stringPoint8);

        stringReference3.Append(formula4);
        stringReference3.Append(stringCache3);

        seriesText2.Append(stringReference3);

        ChartShapeProperties chartShapeProperties5 = new ChartShapeProperties();

        A.SolidFill solidFill5 = new A.SolidFill();
        A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };

        solidFill5.Append(schemeColor5);

        A.Outline outline5 = new A.Outline();
        A.NoFill noFill6 = new A.NoFill();

        outline5.Append(noFill6);
        A.EffectList effectList5 = new A.EffectList();

        chartShapeProperties5.Append(solidFill5);
        chartShapeProperties5.Append(outline5);
        chartShapeProperties5.Append(effectList5);
        InvertIfNegative invertIfNegative2 = new InvertIfNegative() { Val = false };

        DataLabels dataLabels2 = new DataLabels();

        ChartShapeProperties chartShapeProperties6 = new ChartShapeProperties();
        A.NoFill noFill7 = new A.NoFill();

        A.Outline outline6 = new A.Outline();
        A.NoFill noFill8 = new A.NoFill();

        outline6.Append(noFill8);
        A.EffectList effectList6 = new A.EffectList();

        chartShapeProperties6.Append(noFill7);
        chartShapeProperties6.Append(outline6);
        chartShapeProperties6.Append(effectList6);

        TextProperties textProperties3 = new TextProperties();

        A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
        A.ShapeAutoFit shapeAutoFit2 = new A.ShapeAutoFit();

        bodyProperties3.Append(shapeAutoFit2);
        A.ListStyle listStyle3 = new A.ListStyle();

        A.Paragraph paragraph3 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

        A.SolidFill solidFill6 = new A.SolidFill();

        A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 75000 };
        A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 25000 };

        schemeColor6.Append(luminanceModulation4);
        schemeColor6.Append(luminanceOffset4);

        solidFill6.Append(schemeColor6);
        A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties3.Append(solidFill6);
        defaultRunProperties3.Append(latinFont3);
        defaultRunProperties3.Append(eastAsianFont3);
        defaultRunProperties3.Append(complexScriptFont3);

        paragraphProperties3.Append(defaultRunProperties3);
        A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph3.Append(paragraphProperties3);
        paragraph3.Append(endParagraphRunProperties3);

        textProperties3.Append(bodyProperties3);
        textProperties3.Append(listStyle3);
        textProperties3.Append(paragraph3);
        ShowLegendKey showLegendKey2 = new ShowLegendKey() { Val = false };
        ShowValue showValue2 = new ShowValue() { Val = true };
        ShowCategoryName showCategoryName2 = new ShowCategoryName() { Val = false };
        ShowSeriesName showSeriesName2 = new ShowSeriesName() { Val = false };
        ShowPercent showPercent2 = new ShowPercent() { Val = false };
        ShowBubbleSize showBubbleSize2 = new ShowBubbleSize() { Val = false };
        ShowLeaderLines showLeaderLines3 = new ShowLeaderLines() { Val = false };

        DLblsExtensionList dLblsExtensionList2 = new DLblsExtensionList();

        DLblsExtension dLblsExtension2 = new DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
        dLblsExtension2.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
        C15.ShowLeaderLines showLeaderLines4 = new C15.ShowLeaderLines() { Val = true };

        C15.LeaderLines leaderLines2 = new C15.LeaderLines();

        ChartShapeProperties chartShapeProperties7 = new ChartShapeProperties();

        A.Outline outline7 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill7 = new A.SolidFill();

        A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 35000 };
        A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 65000 };

        schemeColor7.Append(luminanceModulation5);
        schemeColor7.Append(luminanceOffset5);

        solidFill7.Append(schemeColor7);
        A.Round round2 = new A.Round();

        outline7.Append(solidFill7);
        outline7.Append(round2);
        A.EffectList effectList7 = new A.EffectList();

        chartShapeProperties7.Append(outline7);
        chartShapeProperties7.Append(effectList7);

        leaderLines2.Append(chartShapeProperties7);

        dLblsExtension2.Append(showLeaderLines4);
        dLblsExtension2.Append(leaderLines2);

        dLblsExtensionList2.Append(dLblsExtension2);

        dataLabels2.Append(chartShapeProperties6);
        dataLabels2.Append(textProperties3);
        dataLabels2.Append(showLegendKey2);
        dataLabels2.Append(showValue2);
        dataLabels2.Append(showCategoryName2);
        dataLabels2.Append(showSeriesName2);
        dataLabels2.Append(showPercent2);
        dataLabels2.Append(showBubbleSize2);
        dataLabels2.Append(showLeaderLines3);
        dataLabels2.Append(dLblsExtensionList2);

        CategoryAxisData categoryAxisData2 = new CategoryAxisData();

        StringReference stringReference4 = new StringReference();
        Formula formula5 = new Formula();
        formula5.Text = "Sheet1!$A$2:$A$7";

        StringCache stringCache4 = new StringCache();
        PointCount pointCount5 = new PointCount() { Val = (UInt32Value)6U };

        StringPoint stringPoint9 = new StringPoint() { Index = (UInt32Value)0U };
        NumericValue numericValue15 = new NumericValue();
        numericValue15.Text = "Schedule 1.0";

        stringPoint9.Append(numericValue15);

        StringPoint stringPoint10 = new StringPoint() { Index = (UInt32Value)1U };
        NumericValue numericValue16 = new NumericValue();
        numericValue16.Text = "Schedule 2.0";

        stringPoint10.Append(numericValue16);

        StringPoint stringPoint11 = new StringPoint() { Index = (UInt32Value)2U };
        NumericValue numericValue17 = new NumericValue();
        numericValue17.Text = "Schedule 3.0";

        stringPoint11.Append(numericValue17);

        StringPoint stringPoint12 = new StringPoint() { Index = (UInt32Value)3U };
        NumericValue numericValue18 = new NumericValue();
        numericValue18.Text = "Schedule 4.0";

        stringPoint12.Append(numericValue18);

        StringPoint stringPoint13 = new StringPoint() { Index = (UInt32Value)4U };
        NumericValue numericValue19 = new NumericValue();
        numericValue19.Text = "Schedule 5.0";

        stringPoint13.Append(numericValue19);

        StringPoint stringPoint14 = new StringPoint() { Index = (UInt32Value)5U };
        NumericValue numericValue20 = new NumericValue();
        numericValue20.Text = "Schedule 6.0";

        stringPoint14.Append(numericValue20);

        stringCache4.Append(pointCount5);
        stringCache4.Append(stringPoint9);
        stringCache4.Append(stringPoint10);
        stringCache4.Append(stringPoint11);
        stringCache4.Append(stringPoint12);
        stringCache4.Append(stringPoint13);
        stringCache4.Append(stringPoint14);

        stringReference4.Append(formula5);
        stringReference4.Append(stringCache4);

        categoryAxisData2.Append(stringReference4);

        Values values2 = new Values();

        NumberReference numberReference2 = new NumberReference();
        Formula formula6 = new Formula();
        formula6.Text = "Sheet1!$C$2:$C$7";

        NumberingCache numberingCache2 = new NumberingCache();
        FormatCode formatCode2 = new FormatCode();
        formatCode2.Text = "General";
        PointCount pointCount6 = new PointCount() { Val = (UInt32Value)6U };

        NumericPoint numericPoint7 = new NumericPoint() { Index = (UInt32Value)0U };
        NumericValue numericValue21 = new NumericValue();
        numericValue21.Text = "2.4";

        numericPoint7.Append(numericValue21);

        NumericPoint numericPoint8 = new NumericPoint() { Index = (UInt32Value)1U };
        NumericValue numericValue22 = new NumericValue();
        numericValue22.Text = "4.4000000000000004";

        numericPoint8.Append(numericValue22);

        NumericPoint numericPoint9 = new NumericPoint() { Index = (UInt32Value)2U };
        NumericValue numericValue23 = new NumericValue();
        numericValue23.Text = "1.8";

        numericPoint9.Append(numericValue23);

        NumericPoint numericPoint10 = new NumericPoint() { Index = (UInt32Value)3U };
        NumericValue numericValue24 = new NumericValue();
        numericValue24.Text = "2.8";

        numericPoint10.Append(numericValue24);

        NumericPoint numericPoint11 = new NumericPoint() { Index = (UInt32Value)4U };
        NumericValue numericValue25 = new NumericValue();
        numericValue25.Text = "2.8";

        numericPoint11.Append(numericValue25);

        NumericPoint numericPoint12 = new NumericPoint() { Index = (UInt32Value)5U };
        NumericValue numericValue26 = new NumericValue();
        numericValue26.Text = "2.8";

        numericPoint12.Append(numericValue26);

        numberingCache2.Append(formatCode2);
        numberingCache2.Append(pointCount6);
        numberingCache2.Append(numericPoint7);
        numberingCache2.Append(numericPoint8);
        numberingCache2.Append(numericPoint9);
        numberingCache2.Append(numericPoint10);
        numberingCache2.Append(numericPoint11);
        numberingCache2.Append(numericPoint12);

        numberReference2.Append(formula6);
        numberReference2.Append(numberingCache2);

        values2.Append(numberReference2);

        BarSerExtensionList barSerExtensionList2 = new BarSerExtensionList();

        BarSerExtension barSerExtension2 = new BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000001-EC2A-431D-B866-E1E729B46B4D}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

        barSerExtension2.Append(openXmlUnknownElement2);

        barSerExtensionList2.Append(barSerExtension2);

        barChartSeries2.Append(index2);
        barChartSeries2.Append(order2);
        barChartSeries2.Append(seriesText2);
        barChartSeries2.Append(chartShapeProperties5);
        barChartSeries2.Append(invertIfNegative2);
        barChartSeries2.Append(dataLabels2);
        barChartSeries2.Append(categoryAxisData2);
        barChartSeries2.Append(values2);
        barChartSeries2.Append(barSerExtensionList2);

        DataLabels dataLabels3 = new DataLabels();
        ShowLegendKey showLegendKey3 = new ShowLegendKey() { Val = false };
        ShowValue showValue3 = new ShowValue() { Val = false };
        ShowCategoryName showCategoryName3 = new ShowCategoryName() { Val = false };
        ShowSeriesName showSeriesName3 = new ShowSeriesName() { Val = false };
        ShowPercent showPercent3 = new ShowPercent() { Val = false };
        ShowBubbleSize showBubbleSize3 = new ShowBubbleSize() { Val = false };

        dataLabels3.Append(showLegendKey3);
        dataLabels3.Append(showValue3);
        dataLabels3.Append(showCategoryName3);
        dataLabels3.Append(showSeriesName3);
        dataLabels3.Append(showPercent3);
        dataLabels3.Append(showBubbleSize3);
        GapWidth gapWidth1 = new GapWidth() { Val = (UInt16Value)150U };
        Overlap overlap1 = new Overlap() { Val = 100 };
        AxisId axisId1 = new AxisId() { Val = (UInt32Value)1956592255U };
        AxisId axisId2 = new AxisId() { Val = (UInt32Value)1956507647U };

        barChart1.Append(barDirection1);
        barChart1.Append(barGrouping1);
        barChart1.Append(varyColors1);
        barChart1.Append(barChartSeries1);
        barChart1.Append(barChartSeries2);
        barChart1.Append(dataLabels3);
        barChart1.Append(gapWidth1);
        barChart1.Append(overlap1);
        barChart1.Append(axisId1);
        barChart1.Append(axisId2);

        CategoryAxis categoryAxis1 = new CategoryAxis();
        AxisId axisId3 = new AxisId() { Val = (UInt32Value)1956592255U };

        Scaling scaling1 = new Scaling();
        Orientation orientation1 = new Orientation() { Val = OrientationValues.MinMax };

        scaling1.Append(orientation1);
        Delete delete1 = new Delete() { Val = false };
        AxisPosition axisPosition1 = new AxisPosition() { Val = AxisPositionValues.Bottom };
        NumberingFormat numberingFormat1 = new NumberingFormat() { FormatCode = "General", SourceLinked = true };
        MajorTickMark majorTickMark1 = new MajorTickMark() { Val = TickMarkValues.None };
        MinorTickMark minorTickMark1 = new MinorTickMark() { Val = TickMarkValues.None };
        TickLabelPosition tickLabelPosition1 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };

        ChartShapeProperties chartShapeProperties8 = new ChartShapeProperties();
        A.NoFill noFill9 = new A.NoFill();

        A.Outline outline8 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill8 = new A.SolidFill();

        A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 15000 };
        A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 85000 };

        schemeColor8.Append(luminanceModulation6);
        schemeColor8.Append(luminanceOffset6);

        solidFill8.Append(schemeColor8);
        A.Round round3 = new A.Round();

        outline8.Append(solidFill8);
        outline8.Append(round3);
        A.EffectList effectList8 = new A.EffectList();

        chartShapeProperties8.Append(noFill9);
        chartShapeProperties8.Append(outline8);
        chartShapeProperties8.Append(effectList8);

        TextProperties textProperties4 = new TextProperties();
        A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
        A.ListStyle listStyle4 = new A.ListStyle();

        A.Paragraph paragraph4 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

        A.SolidFill solidFill9 = new A.SolidFill();

        A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor9.Append(luminanceModulation7);
        schemeColor9.Append(luminanceOffset7);

        solidFill9.Append(schemeColor9);
        A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties4.Append(solidFill9);
        defaultRunProperties4.Append(latinFont4);
        defaultRunProperties4.Append(eastAsianFont4);
        defaultRunProperties4.Append(complexScriptFont4);

        paragraphProperties4.Append(defaultRunProperties4);
        A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph4.Append(paragraphProperties4);
        paragraph4.Append(endParagraphRunProperties4);

        textProperties4.Append(bodyProperties4);
        textProperties4.Append(listStyle4);
        textProperties4.Append(paragraph4);
        CrossingAxis crossingAxis1 = new CrossingAxis() { Val = (UInt32Value)1956507647U };
        Crosses crosses1 = new Crosses() { Val = CrossesValues.AutoZero };
        AutoLabeled autoLabeled1 = new AutoLabeled() { Val = true };
        LabelAlignment labelAlignment1 = new LabelAlignment() { Val = LabelAlignmentValues.Center };
        LabelOffset labelOffset1 = new LabelOffset() { Val = (UInt16Value)100U };
        NoMultiLevelLabels noMultiLevelLabels1 = new NoMultiLevelLabels() { Val = false };

        categoryAxis1.Append(axisId3);
        categoryAxis1.Append(scaling1);
        categoryAxis1.Append(delete1);
        categoryAxis1.Append(axisPosition1);
        categoryAxis1.Append(numberingFormat1);
        categoryAxis1.Append(majorTickMark1);
        categoryAxis1.Append(minorTickMark1);
        categoryAxis1.Append(tickLabelPosition1);
        categoryAxis1.Append(chartShapeProperties8);
        categoryAxis1.Append(textProperties4);
        categoryAxis1.Append(crossingAxis1);
        categoryAxis1.Append(crosses1);
        categoryAxis1.Append(autoLabeled1);
        categoryAxis1.Append(labelAlignment1);
        categoryAxis1.Append(labelOffset1);
        categoryAxis1.Append(noMultiLevelLabels1);

        ValueAxis valueAxis1 = new ValueAxis();
        AxisId axisId4 = new AxisId() { Val = (UInt32Value)1956507647U };

        Scaling scaling2 = new Scaling();
        Orientation orientation2 = new Orientation() { Val = OrientationValues.MinMax };

        scaling2.Append(orientation2);
        Delete delete2 = new Delete() { Val = false };
        AxisPosition axisPosition2 = new AxisPosition() { Val = AxisPositionValues.Left };

        MajorGridlines majorGridlines1 = new MajorGridlines();

        ChartShapeProperties chartShapeProperties9 = new ChartShapeProperties();

        A.Outline outline9 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill10 = new A.SolidFill();

        A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 15000 };
        A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 85000 };

        schemeColor10.Append(luminanceModulation8);
        schemeColor10.Append(luminanceOffset8);

        solidFill10.Append(schemeColor10);
        A.Round round4 = new A.Round();

        outline9.Append(solidFill10);
        outline9.Append(round4);
        A.EffectList effectList9 = new A.EffectList();

        chartShapeProperties9.Append(outline9);
        chartShapeProperties9.Append(effectList9);

        majorGridlines1.Append(chartShapeProperties9);
        NumberingFormat numberingFormat2 = new NumberingFormat() { FormatCode = "General", SourceLinked = true };
        MajorTickMark majorTickMark2 = new MajorTickMark() { Val = TickMarkValues.None };
        MinorTickMark minorTickMark2 = new MinorTickMark() { Val = TickMarkValues.None };
        TickLabelPosition tickLabelPosition2 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };

        ChartShapeProperties chartShapeProperties10 = new ChartShapeProperties();
        A.NoFill noFill10 = new A.NoFill();

        A.Outline outline10 = new A.Outline();
        A.NoFill noFill11 = new A.NoFill();

        outline10.Append(noFill11);
        A.EffectList effectList10 = new A.EffectList();

        chartShapeProperties10.Append(noFill10);
        chartShapeProperties10.Append(outline10);
        chartShapeProperties10.Append(effectList10);

        TextProperties textProperties5 = new TextProperties();
        A.BodyProperties bodyProperties5 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
        A.ListStyle listStyle5 = new A.ListStyle();

        A.Paragraph paragraph5 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

        A.SolidFill solidFill11 = new A.SolidFill();

        A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor11.Append(luminanceModulation9);
        schemeColor11.Append(luminanceOffset9);

        solidFill11.Append(schemeColor11);
        A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties5.Append(solidFill11);
        defaultRunProperties5.Append(latinFont5);
        defaultRunProperties5.Append(eastAsianFont5);
        defaultRunProperties5.Append(complexScriptFont5);

        paragraphProperties5.Append(defaultRunProperties5);
        A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph5.Append(paragraphProperties5);
        paragraph5.Append(endParagraphRunProperties5);

        textProperties5.Append(bodyProperties5);
        textProperties5.Append(listStyle5);
        textProperties5.Append(paragraph5);
        CrossingAxis crossingAxis2 = new CrossingAxis() { Val = (UInt32Value)1956592255U };
        Crosses crosses2 = new Crosses() { Val = CrossesValues.AutoZero };
        CrossBetween crossBetween1 = new CrossBetween() { Val = CrossBetweenValues.Between };

        valueAxis1.Append(axisId4);
        valueAxis1.Append(scaling2);
        valueAxis1.Append(delete2);
        valueAxis1.Append(axisPosition2);
        valueAxis1.Append(majorGridlines1);
        valueAxis1.Append(numberingFormat2);
        valueAxis1.Append(majorTickMark2);
        valueAxis1.Append(minorTickMark2);
        valueAxis1.Append(tickLabelPosition2);
        valueAxis1.Append(chartShapeProperties10);
        valueAxis1.Append(textProperties5);
        valueAxis1.Append(crossingAxis2);
        valueAxis1.Append(crosses2);
        valueAxis1.Append(crossBetween1);

        ShapeProperties shapeProperties1 = new ShapeProperties();
        A.NoFill noFill12 = new A.NoFill();

        A.Outline outline11 = new A.Outline();
        A.NoFill noFill13 = new A.NoFill();

        outline11.Append(noFill13);
        A.EffectList effectList11 = new A.EffectList();

        shapeProperties1.Append(noFill12);
        shapeProperties1.Append(outline11);
        shapeProperties1.Append(effectList11);

        plotArea1.Append(layout1);
        plotArea1.Append(barChart1);
        plotArea1.Append(categoryAxis1);
        plotArea1.Append(valueAxis1);
        plotArea1.Append(shapeProperties1);

        Legend legend1 = new Legend();
        LegendPosition legendPosition1 = new LegendPosition() { Val = LegendPositionValues.Bottom };
        Overlay overlay2 = new Overlay() { Val = false };

        ChartShapeProperties chartShapeProperties11 = new ChartShapeProperties();
        A.NoFill noFill14 = new A.NoFill();

        A.Outline outline12 = new A.Outline();
        A.NoFill noFill15 = new A.NoFill();

        outline12.Append(noFill15);
        A.EffectList effectList12 = new A.EffectList();

        chartShapeProperties11.Append(noFill14);
        chartShapeProperties11.Append(outline12);
        chartShapeProperties11.Append(effectList12);

        TextProperties textProperties6 = new TextProperties();
        A.BodyProperties bodyProperties6 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
        A.ListStyle listStyle6 = new A.ListStyle();

        A.Paragraph paragraph6 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

        A.SolidFill solidFill12 = new A.SolidFill();

        A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor12.Append(luminanceModulation10);
        schemeColor12.Append(luminanceOffset10);

        solidFill12.Append(schemeColor12);
        A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties6.Append(solidFill12);
        defaultRunProperties6.Append(latinFont6);
        defaultRunProperties6.Append(eastAsianFont6);
        defaultRunProperties6.Append(complexScriptFont6);

        paragraphProperties6.Append(defaultRunProperties6);
        A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph6.Append(paragraphProperties6);
        paragraph6.Append(endParagraphRunProperties6);

        textProperties6.Append(bodyProperties6);
        textProperties6.Append(listStyle6);
        textProperties6.Append(paragraph6);

        legend1.Append(legendPosition1);
        legend1.Append(overlay2);
        legend1.Append(chartShapeProperties11);
        legend1.Append(textProperties6);
        PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
        DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };

        ExtensionList extensionList1 = new ExtensionList();

        Extension extension1 = new Extension() { Uri = "{56B9EC1D-385E-4148-901F-78D8002777C0}" };
        extension1.AddNamespaceDeclaration("c16r3", "http://schemas.microsoft.com/office/drawing/2017/03/chart");

        OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16r3:dataDisplayOptions16 xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\"><c16r3:dispNaAsBlank val=\"1\" /></c16r3:dataDisplayOptions16>");

        extension1.Append(openXmlUnknownElement3);

        extensionList1.Append(extension1);
        ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };

        chart1.Append(title1);
        chart1.Append(autoTitleDeleted1);
        chart1.Append(plotArea1);
        chart1.Append(legend1);
        chart1.Append(plotVisibleOnly1);
        chart1.Append(displayBlanksAs1);
        chart1.Append(extensionList1);
        chart1.Append(showDataLabelsOverMaximum1);

        ShapeProperties shapeProperties2 = new ShapeProperties();

        A.SolidFill solidFill13 = new A.SolidFill();
        A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

        solidFill13.Append(schemeColor13);

        A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill14 = new A.SolidFill();

        A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 15000 };
        A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 85000 };

        schemeColor14.Append(luminanceModulation11);
        schemeColor14.Append(luminanceOffset11);

        solidFill14.Append(schemeColor14);
        A.Round round5 = new A.Round();

        outline13.Append(solidFill14);
        outline13.Append(round5);
        A.EffectList effectList13 = new A.EffectList();

        shapeProperties2.Append(solidFill13);
        shapeProperties2.Append(outline13);
        shapeProperties2.Append(effectList13);

        TextProperties textProperties7 = new TextProperties();
        A.BodyProperties bodyProperties7 = new A.BodyProperties();
        A.ListStyle listStyle7 = new A.ListStyle();

        A.Paragraph paragraph7 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties();
        A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties();

        paragraphProperties7.Append(defaultRunProperties7);
        A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph7.Append(paragraphProperties7);
        paragraph7.Append(endParagraphRunProperties7);

        textProperties7.Append(bodyProperties7);
        textProperties7.Append(listStyle7);
        textProperties7.Append(paragraph7);

        ExternalData externalData1 = new ExternalData() { Id = "rId3" };
        AutoUpdate autoUpdate1 = new AutoUpdate() { Val = false };

        externalData1.Append(autoUpdate1);

        chartSpace1.Append(date19041);
        chartSpace1.Append(editingLanguage1);
        chartSpace1.Append(roundedCorners1);
        chartSpace1.Append(alternateContent1);
        chartSpace1.Append(chart1);
        chartSpace1.Append(shapeProperties2);
        chartSpace1.Append(textProperties7);
        chartSpace1.Append(externalData1);
        return chartSpace1;
    }

}

{
    ChartSpace chartSpace1 = new ChartSpace();
    chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
    chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
    Date1904 date19041 = new Date1904() { Val = false };
    EditingLanguage editingLanguage1 = new EditingLanguage() { Val = "en-US" };
    RoundedCorners roundedCorners1 = new RoundedCorners() { Val = false };

    AlternateContent alternateContent1 = new AlternateContent();
    alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

    AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
    alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
    C14.Style style1 = new C14.Style() { Val = 102 };

    alternateContentChoice1.Append(style1);

    AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
    Style style2 = new Style() { Val = 2 };

    alternateContentFallback1.Append(style2);

    alternateContent1.Append(alternateContentChoice1);
    alternateContent1.Append(alternateContentFallback1);

    Chart chart1 = new Chart();

    Title title1 = new Title();
    Overlay overlay1 = new Overlay() { Val = false };

    ChartShapeProperties chartShapeProperties1 = new ChartShapeProperties();
    A.NoFill noFill1 = new A.NoFill();

    A.Outline outline1 = new A.Outline();
    A.NoFill noFill2 = new A.NoFill();

    outline1.Append(noFill2);
    A.EffectList effectList1 = new A.EffectList();

    chartShapeProperties1.Append(noFill1);
    chartShapeProperties1.Append(outline1);
    chartShapeProperties1.Append(effectList1);

    TextProperties textProperties1 = new TextProperties();
    A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ListStyle listStyle1 = new A.ListStyle();

    A.Paragraph paragraph1 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };

    A.SolidFill solidFill1 = new A.SolidFill();

    A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 65000 };
    A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

    schemeColor1.Append(luminanceModulation1);
    schemeColor1.Append(luminanceOffset1);

    solidFill1.Append(schemeColor1);
    A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties1.Append(solidFill1);
    defaultRunProperties1.Append(latinFont1);
    defaultRunProperties1.Append(eastAsianFont1);
    defaultRunProperties1.Append(complexScriptFont1);

    paragraphProperties1.Append(defaultRunProperties1);
    A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph1.Append(paragraphProperties1);
    paragraph1.Append(endParagraphRunProperties1);

    textProperties1.Append(bodyProperties1);
    textProperties1.Append(listStyle1);
    textProperties1.Append(paragraph1);

    title1.Append(overlay1);
    title1.Append(chartShapeProperties1);
    title1.Append(textProperties1);
    AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };

    PlotArea plotArea1 = new PlotArea();
    Layout layout1 = new Layout();

    BarChart barChart1 = new BarChart();
    BarDirection barDirection1 = new BarDirection() { Val = BarDirectionValues.Column };
    BarGrouping barGrouping1 = new BarGrouping() { Val = BarGroupingValues.Stacked };
    VaryColors varyColors1 = new VaryColors() { Val = false };

    BarChartSeries barChartSeries1 = new BarChartSeries();
    Index index1 = new Index() { Val = (UInt32Value)0U };
    Order order1 = new Order() { Val = (UInt32Value)0U };

    SeriesText seriesText1 = new SeriesText();

    StringReference stringReference1 = new StringReference();
    Formula formula1 = new Formula();
    formula1.Text = "Sheet1!$B$1";

    StringCache stringCache1 = new StringCache();
    PointCount pointCount1 = new PointCount() { Val = (UInt32Value)1U };

    StringPoint stringPoint1 = new StringPoint() { Index = (UInt32Value)0U };
    NumericValue numericValue1 = new NumericValue();
    numericValue1.Text = "Claimed to Date";

    stringPoint1.Append(numericValue1);

    stringCache1.Append(pointCount1);
    stringCache1.Append(stringPoint1);

    stringReference1.Append(formula1);
    stringReference1.Append(stringCache1);

    seriesText1.Append(stringReference1);

    ChartShapeProperties chartShapeProperties2 = new ChartShapeProperties();

    A.SolidFill solidFill2 = new A.SolidFill();
    A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

    solidFill2.Append(schemeColor2);

    A.Outline outline2 = new A.Outline();
    A.NoFill noFill3 = new A.NoFill();

    outline2.Append(noFill3);
    A.EffectList effectList2 = new A.EffectList();

    chartShapeProperties2.Append(solidFill2);
    chartShapeProperties2.Append(outline2);
    chartShapeProperties2.Append(effectList2);
    InvertIfNegative invertIfNegative1 = new InvertIfNegative() { Val = false };

    DataLabels dataLabels1 = new DataLabels();

    ChartShapeProperties chartShapeProperties3 = new ChartShapeProperties();
    A.NoFill noFill4 = new A.NoFill();

    A.Outline outline3 = new A.Outline();
    A.NoFill noFill5 = new A.NoFill();

    outline3.Append(noFill5);
    A.EffectList effectList3 = new A.EffectList();

    chartShapeProperties3.Append(noFill4);
    chartShapeProperties3.Append(outline3);
    chartShapeProperties3.Append(effectList3);

    TextProperties textProperties2 = new TextProperties();

    A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

    bodyProperties2.Append(shapeAutoFit1);
    A.ListStyle listStyle2 = new A.ListStyle();

    A.Paragraph paragraph2 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill3 = new A.SolidFill();

    A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 75000 };
    A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 25000 };

    schemeColor3.Append(luminanceModulation2);
    schemeColor3.Append(luminanceOffset2);

    solidFill3.Append(schemeColor3);
    A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties2.Append(solidFill3);
    defaultRunProperties2.Append(latinFont2);
    defaultRunProperties2.Append(eastAsianFont2);
    defaultRunProperties2.Append(complexScriptFont2);

    paragraphProperties2.Append(defaultRunProperties2);
    A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph2.Append(paragraphProperties2);
    paragraph2.Append(endParagraphRunProperties2);

    textProperties2.Append(bodyProperties2);
    textProperties2.Append(listStyle2);
    textProperties2.Append(paragraph2);
    ShowLegendKey showLegendKey1 = new ShowLegendKey() { Val = false };
    ShowValue showValue1 = new ShowValue() { Val = true };
    ShowCategoryName showCategoryName1 = new ShowCategoryName() { Val = false };
    ShowSeriesName showSeriesName1 = new ShowSeriesName() { Val = false };
    ShowPercent showPercent1 = new ShowPercent() { Val = false };
    ShowBubbleSize showBubbleSize1 = new ShowBubbleSize() { Val = false };
    ShowLeaderLines showLeaderLines1 = new ShowLeaderLines() { Val = false };

    DLblsExtensionList dLblsExtensionList1 = new DLblsExtensionList();

    DLblsExtension dLblsExtension1 = new DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
    dLblsExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
    C15.ShowLeaderLines showLeaderLines2 = new C15.ShowLeaderLines() { Val = true };

    C15.LeaderLines leaderLines1 = new C15.LeaderLines();

    ChartShapeProperties chartShapeProperties4 = new ChartShapeProperties();

    A.Outline outline4 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill4 = new A.SolidFill();

    A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 35000 };
    A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 65000 };

    schemeColor4.Append(luminanceModulation3);
    schemeColor4.Append(luminanceOffset3);

    solidFill4.Append(schemeColor4);
    A.Round round1 = new A.Round();

    outline4.Append(solidFill4);
    outline4.Append(round1);
    A.EffectList effectList4 = new A.EffectList();

    chartShapeProperties4.Append(outline4);
    chartShapeProperties4.Append(effectList4);

    leaderLines1.Append(chartShapeProperties4);

    dLblsExtension1.Append(showLeaderLines2);
    dLblsExtension1.Append(leaderLines1);

    dLblsExtensionList1.Append(dLblsExtension1);

    dataLabels1.Append(chartShapeProperties3);
    dataLabels1.Append(textProperties2);
    dataLabels1.Append(showLegendKey1);
    dataLabels1.Append(showValue1);
    dataLabels1.Append(showCategoryName1);
    dataLabels1.Append(showSeriesName1);
    dataLabels1.Append(showPercent1);
    dataLabels1.Append(showBubbleSize1);
    dataLabels1.Append(showLeaderLines1);
    dataLabels1.Append(dLblsExtensionList1);

    CategoryAxisData categoryAxisData1 = new CategoryAxisData();

    StringReference stringReference2 = new StringReference();
    Formula formula2 = new Formula();
    formula2.Text = "Sheet1!$A$2:$A$7";

    StringCache stringCache2 = new StringCache();
    PointCount pointCount2 = new PointCount() { Val = (UInt32Value)6U };

    StringPoint stringPoint2 = new StringPoint() { Index = (UInt32Value)0U };
    NumericValue numericValue2 = new NumericValue();
    numericValue2.Text = "Schedule 1.0";

    stringPoint2.Append(numericValue2);

    StringPoint stringPoint3 = new StringPoint() { Index = (UInt32Value)1U };
    NumericValue numericValue3 = new NumericValue();
    numericValue3.Text = "Schedule 2.0";

    stringPoint3.Append(numericValue3);

    StringPoint stringPoint4 = new StringPoint() { Index = (UInt32Value)2U };
    NumericValue numericValue4 = new NumericValue();
    numericValue4.Text = "Schedule 3.0";

    stringPoint4.Append(numericValue4);

    StringPoint stringPoint5 = new StringPoint() { Index = (UInt32Value)3U };
    NumericValue numericValue5 = new NumericValue();
    numericValue5.Text = "Schedule 4.0";

    stringPoint5.Append(numericValue5);

    StringPoint stringPoint6 = new StringPoint() { Index = (UInt32Value)4U };
    NumericValue numericValue6 = new NumericValue();
    numericValue6.Text = "Schedule 5.0";

    stringPoint6.Append(numericValue6);

    StringPoint stringPoint7 = new StringPoint() { Index = (UInt32Value)5U };
    NumericValue numericValue7 = new NumericValue();
    numericValue7.Text = "Schedule 6.0";

    stringPoint7.Append(numericValue7);

    stringCache2.Append(pointCount2);
    stringCache2.Append(stringPoint2);
    stringCache2.Append(stringPoint3);
    stringCache2.Append(stringPoint4);
    stringCache2.Append(stringPoint5);
    stringCache2.Append(stringPoint6);
    stringCache2.Append(stringPoint7);

    stringReference2.Append(formula2);
    stringReference2.Append(stringCache2);

    categoryAxisData1.Append(stringReference2);

    Values values1 = new Values();

    NumberReference numberReference1 = new NumberReference();
    Formula formula3 = new Formula();
    formula3.Text = "Sheet1!$B$2:$B$7";

    NumberingCache numberingCache1 = new NumberingCache();
    FormatCode formatCode1 = new FormatCode();
    formatCode1.Text = "General";
    PointCount pointCount3 = new PointCount() { Val = (UInt32Value)6U };

    NumericPoint numericPoint1 = new NumericPoint() { Index = (UInt32Value)0U };
    NumericValue numericValue8 = new NumericValue();
    numericValue8.Text = "4.3";

    numericPoint1.Append(numericValue8);

    NumericPoint numericPoint2 = new NumericPoint() { Index = (UInt32Value)1U };
    NumericValue numericValue9 = new NumericValue();
    numericValue9.Text = "2.5";

    numericPoint2.Append(numericValue9);

    NumericPoint numericPoint3 = new NumericPoint() { Index = (UInt32Value)2U };
    NumericValue numericValue10 = new NumericValue();
    numericValue10.Text = "3.5";

    numericPoint3.Append(numericValue10);

    NumericPoint numericPoint4 = new NumericPoint() { Index = (UInt32Value)3U };
    NumericValue numericValue11 = new NumericValue();
    numericValue11.Text = "4.5";

    numericPoint4.Append(numericValue11);

    NumericPoint numericPoint5 = new NumericPoint() { Index = (UInt32Value)4U };
    NumericValue numericValue12 = new NumericValue();
    numericValue12.Text = "4.5";

    numericPoint5.Append(numericValue12);

    NumericPoint numericPoint6 = new NumericPoint() { Index = (UInt32Value)5U };
    NumericValue numericValue13 = new NumericValue();
    numericValue13.Text = "4.5";

    numericPoint6.Append(numericValue13);

    numberingCache1.Append(formatCode1);
    numberingCache1.Append(pointCount3);
    numberingCache1.Append(numericPoint1);
    numberingCache1.Append(numericPoint2);
    numberingCache1.Append(numericPoint3);
    numberingCache1.Append(numericPoint4);
    numberingCache1.Append(numericPoint5);
    numberingCache1.Append(numericPoint6);

    numberReference1.Append(formula3);
    numberReference1.Append(numberingCache1);

    values1.Append(numberReference1);

    BarSerExtensionList barSerExtensionList1 = new BarSerExtensionList();

    BarSerExtension barSerExtension1 = new BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
    barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

    OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-EC2A-431D-B866-E1E729B46B4D}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

    barSerExtension1.Append(openXmlUnknownElement1);

    barSerExtensionList1.Append(barSerExtension1);

    barChartSeries1.Append(index1);
    barChartSeries1.Append(order1);
    barChartSeries1.Append(seriesText1);
    barChartSeries1.Append(chartShapeProperties2);
    barChartSeries1.Append(invertIfNegative1);
    barChartSeries1.Append(dataLabels1);
    barChartSeries1.Append(categoryAxisData1);
    barChartSeries1.Append(values1);
    barChartSeries1.Append(barSerExtensionList1);

    BarChartSeries barChartSeries2 = new BarChartSeries();
    Index index2 = new Index() { Val = (UInt32Value)1U };
    Order order2 = new Order() { Val = (UInt32Value)1U };

    SeriesText seriesText2 = new SeriesText();

    StringReference stringReference3 = new StringReference();
    Formula formula4 = new Formula();
    formula4.Text = "Sheet1!$C$1";

    StringCache stringCache3 = new StringCache();
    PointCount pointCount4 = new PointCount() { Val = (UInt32Value)1U };

    StringPoint stringPoint8 = new StringPoint() { Index = (UInt32Value)0U };
    NumericValue numericValue14 = new NumericValue();
    numericValue14.Text = "Contract Work to Claim";

    stringPoint8.Append(numericValue14);

    stringCache3.Append(pointCount4);
    stringCache3.Append(stringPoint8);

    stringReference3.Append(formula4);
    stringReference3.Append(stringCache3);

    seriesText2.Append(stringReference3);

    ChartShapeProperties chartShapeProperties5 = new ChartShapeProperties();

    A.SolidFill solidFill5 = new A.SolidFill();
    A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };

    solidFill5.Append(schemeColor5);

    A.Outline outline5 = new A.Outline();
    A.NoFill noFill6 = new A.NoFill();

    outline5.Append(noFill6);
    A.EffectList effectList5 = new A.EffectList();

    chartShapeProperties5.Append(solidFill5);
    chartShapeProperties5.Append(outline5);
    chartShapeProperties5.Append(effectList5);
    InvertIfNegative invertIfNegative2 = new InvertIfNegative() { Val = false };

    DataLabels dataLabels2 = new DataLabels();

    ChartShapeProperties chartShapeProperties6 = new ChartShapeProperties();
    A.NoFill noFill7 = new A.NoFill();

    A.Outline outline6 = new A.Outline();
    A.NoFill noFill8 = new A.NoFill();

    outline6.Append(noFill8);
    A.EffectList effectList6 = new A.EffectList();

    chartShapeProperties6.Append(noFill7);
    chartShapeProperties6.Append(outline6);
    chartShapeProperties6.Append(effectList6);

    TextProperties textProperties3 = new TextProperties();

    A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ShapeAutoFit shapeAutoFit2 = new A.ShapeAutoFit();

    bodyProperties3.Append(shapeAutoFit2);
    A.ListStyle listStyle3 = new A.ListStyle();

    A.Paragraph paragraph3 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill6 = new A.SolidFill();

    A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 75000 };
    A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 25000 };

    schemeColor6.Append(luminanceModulation4);
    schemeColor6.Append(luminanceOffset4);

    solidFill6.Append(schemeColor6);
    A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties3.Append(solidFill6);
    defaultRunProperties3.Append(latinFont3);
    defaultRunProperties3.Append(eastAsianFont3);
    defaultRunProperties3.Append(complexScriptFont3);

    paragraphProperties3.Append(defaultRunProperties3);
    A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph3.Append(paragraphProperties3);
    paragraph3.Append(endParagraphRunProperties3);

    textProperties3.Append(bodyProperties3);
    textProperties3.Append(listStyle3);
    textProperties3.Append(paragraph3);
    ShowLegendKey showLegendKey2 = new ShowLegendKey() { Val = false };
    ShowValue showValue2 = new ShowValue() { Val = true };
    ShowCategoryName showCategoryName2 = new ShowCategoryName() { Val = false };
    ShowSeriesName showSeriesName2 = new ShowSeriesName() { Val = false };
    ShowPercent showPercent2 = new ShowPercent() { Val = false };
    ShowBubbleSize showBubbleSize2 = new ShowBubbleSize() { Val = false };
    ShowLeaderLines showLeaderLines3 = new ShowLeaderLines() { Val = false };

    DLblsExtensionList dLblsExtensionList2 = new DLblsExtensionList();

    DLblsExtension dLblsExtension2 = new DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
    dLblsExtension2.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
    C15.ShowLeaderLines showLeaderLines4 = new C15.ShowLeaderLines() { Val = true };

    C15.LeaderLines leaderLines2 = new C15.LeaderLines();

    ChartShapeProperties chartShapeProperties7 = new ChartShapeProperties();

    A.Outline outline7 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill7 = new A.SolidFill();

    A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 35000 };
    A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 65000 };

    schemeColor7.Append(luminanceModulation5);
    schemeColor7.Append(luminanceOffset5);

    solidFill7.Append(schemeColor7);
    A.Round round2 = new A.Round();

    outline7.Append(solidFill7);
    outline7.Append(round2);
    A.EffectList effectList7 = new A.EffectList();

    chartShapeProperties7.Append(outline7);
    chartShapeProperties7.Append(effectList7);

    leaderLines2.Append(chartShapeProperties7);

    dLblsExtension2.Append(showLeaderLines4);
    dLblsExtension2.Append(leaderLines2);

    dLblsExtensionList2.Append(dLblsExtension2);

    dataLabels2.Append(chartShapeProperties6);
    dataLabels2.Append(textProperties3);
    dataLabels2.Append(showLegendKey2);
    dataLabels2.Append(showValue2);
    dataLabels2.Append(showCategoryName2);
    dataLabels2.Append(showSeriesName2);
    dataLabels2.Append(showPercent2);
    dataLabels2.Append(showBubbleSize2);
    dataLabels2.Append(showLeaderLines3);
    dataLabels2.Append(dLblsExtensionList2);

    CategoryAxisData categoryAxisData2 = new CategoryAxisData();

    StringReference stringReference4 = new StringReference();
    Formula formula5 = new Formula();
    formula5.Text = "Sheet1!$A$2:$A$7";

    StringCache stringCache4 = new StringCache();
    PointCount pointCount5 = new PointCount() { Val = (UInt32Value)6U };

    StringPoint stringPoint9 = new StringPoint() { Index = (UInt32Value)0U };
    NumericValue numericValue15 = new NumericValue();
    numericValue15.Text = "Schedule 1.0";

    stringPoint9.Append(numericValue15);

    StringPoint stringPoint10 = new StringPoint() { Index = (UInt32Value)1U };
    NumericValue numericValue16 = new NumericValue();
    numericValue16.Text = "Schedule 2.0";

    stringPoint10.Append(numericValue16);

    StringPoint stringPoint11 = new StringPoint() { Index = (UInt32Value)2U };
    NumericValue numericValue17 = new NumericValue();
    numericValue17.Text = "Schedule 3.0";

    stringPoint11.Append(numericValue17);

    StringPoint stringPoint12 = new StringPoint() { Index = (UInt32Value)3U };
    NumericValue numericValue18 = new NumericValue();
    numericValue18.Text = "Schedule 4.0";

    stringPoint12.Append(numericValue18);

    StringPoint stringPoint13 = new StringPoint() { Index = (UInt32Value)4U };
    NumericValue numericValue19 = new NumericValue();
    numericValue19.Text = "Schedule 5.0";

    stringPoint13.Append(numericValue19);

    StringPoint stringPoint14 = new StringPoint() { Index = (UInt32Value)5U };
    NumericValue numericValue20 = new NumericValue();
    numericValue20.Text = "Schedule 6.0";

    stringPoint14.Append(numericValue20);

    stringCache4.Append(pointCount5);
    stringCache4.Append(stringPoint9);
    stringCache4.Append(stringPoint10);
    stringCache4.Append(stringPoint11);
    stringCache4.Append(stringPoint12);
    stringCache4.Append(stringPoint13);
    stringCache4.Append(stringPoint14);

    stringReference4.Append(formula5);
    stringReference4.Append(stringCache4);

    categoryAxisData2.Append(stringReference4);

    Values values2 = new Values();

    NumberReference numberReference2 = new NumberReference();
    Formula formula6 = new Formula();
    formula6.Text = "Sheet1!$C$2:$C$7";

    NumberingCache numberingCache2 = new NumberingCache();
    FormatCode formatCode2 = new FormatCode();
    formatCode2.Text = "General";
    PointCount pointCount6 = new PointCount() { Val = (UInt32Value)6U };

    NumericPoint numericPoint7 = new NumericPoint() { Index = (UInt32Value)0U };
    NumericValue numericValue21 = new NumericValue();
    numericValue21.Text = "2.4";

    numericPoint7.Append(numericValue21);

    NumericPoint numericPoint8 = new NumericPoint() { Index = (UInt32Value)1U };
    NumericValue numericValue22 = new NumericValue();
    numericValue22.Text = "4.4000000000000004";

    numericPoint8.Append(numericValue22);

    NumericPoint numericPoint9 = new NumericPoint() { Index = (UInt32Value)2U };
    NumericValue numericValue23 = new NumericValue();
    numericValue23.Text = "1.8";

    numericPoint9.Append(numericValue23);

    NumericPoint numericPoint10 = new NumericPoint() { Index = (UInt32Value)3U };
    NumericValue numericValue24 = new NumericValue();
    numericValue24.Text = "2.8";

    numericPoint10.Append(numericValue24);

    NumericPoint numericPoint11 = new NumericPoint() { Index = (UInt32Value)4U };
    NumericValue numericValue25 = new NumericValue();
    numericValue25.Text = "2.8";

    numericPoint11.Append(numericValue25);

    NumericPoint numericPoint12 = new NumericPoint() { Index = (UInt32Value)5U };
    NumericValue numericValue26 = new NumericValue();
    numericValue26.Text = "2.8";

    numericPoint12.Append(numericValue26);

    numberingCache2.Append(formatCode2);
    numberingCache2.Append(pointCount6);
    numberingCache2.Append(numericPoint7);
    numberingCache2.Append(numericPoint8);
    numberingCache2.Append(numericPoint9);
    numberingCache2.Append(numericPoint10);
    numberingCache2.Append(numericPoint11);
    numberingCache2.Append(numericPoint12);

    numberReference2.Append(formula6);
    numberReference2.Append(numberingCache2);

    values2.Append(numberReference2);

    BarSerExtensionList barSerExtensionList2 = new BarSerExtensionList();

    BarSerExtension barSerExtension2 = new BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
    barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

    OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000001-EC2A-431D-B866-E1E729B46B4D}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

    barSerExtension2.Append(openXmlUnknownElement2);

    barSerExtensionList2.Append(barSerExtension2);

    barChartSeries2.Append(index2);
    barChartSeries2.Append(order2);
    barChartSeries2.Append(seriesText2);
    barChartSeries2.Append(chartShapeProperties5);
    barChartSeries2.Append(invertIfNegative2);
    barChartSeries2.Append(dataLabels2);
    barChartSeries2.Append(categoryAxisData2);
    barChartSeries2.Append(values2);
    barChartSeries2.Append(barSerExtensionList2);

    DataLabels dataLabels3 = new DataLabels();
    ShowLegendKey showLegendKey3 = new ShowLegendKey() { Val = false };
    ShowValue showValue3 = new ShowValue() { Val = false };
    ShowCategoryName showCategoryName3 = new ShowCategoryName() { Val = false };
    ShowSeriesName showSeriesName3 = new ShowSeriesName() { Val = false };
    ShowPercent showPercent3 = new ShowPercent() { Val = false };
    ShowBubbleSize showBubbleSize3 = new ShowBubbleSize() { Val = false };

    dataLabels3.Append(showLegendKey3);
    dataLabels3.Append(showValue3);
    dataLabels3.Append(showCategoryName3);
    dataLabels3.Append(showSeriesName3);
    dataLabels3.Append(showPercent3);
    dataLabels3.Append(showBubbleSize3);
    GapWidth gapWidth1 = new GapWidth() { Val = (UInt16Value)150U };
    Overlap overlap1 = new Overlap() { Val = 100 };
    AxisId axisId1 = new AxisId() { Val = (UInt32Value)1956592255U };
    AxisId axisId2 = new AxisId() { Val = (UInt32Value)1956507647U };

    barChart1.Append(barDirection1);
    barChart1.Append(barGrouping1);
    barChart1.Append(varyColors1);
    barChart1.Append(barChartSeries1);
    barChart1.Append(barChartSeries2);
    barChart1.Append(dataLabels3);
    barChart1.Append(gapWidth1);
    barChart1.Append(overlap1);
    barChart1.Append(axisId1);
    barChart1.Append(axisId2);

    CategoryAxis categoryAxis1 = new CategoryAxis();
    AxisId axisId3 = new AxisId() { Val = (UInt32Value)1956592255U };

    Scaling scaling1 = new Scaling();
    Orientation orientation1 = new Orientation() { Val = OrientationValues.MinMax };

    scaling1.Append(orientation1);
    Delete delete1 = new Delete() { Val = false };
    AxisPosition axisPosition1 = new AxisPosition() { Val = AxisPositionValues.Bottom };
    NumberingFormat numberingFormat1 = new NumberingFormat() { FormatCode = "General", SourceLinked = true };
    MajorTickMark majorTickMark1 = new MajorTickMark() { Val = TickMarkValues.None };
    MinorTickMark minorTickMark1 = new MinorTickMark() { Val = TickMarkValues.None };
    TickLabelPosition tickLabelPosition1 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };

    ChartShapeProperties chartShapeProperties8 = new ChartShapeProperties();
    A.NoFill noFill9 = new A.NoFill();

    A.Outline outline8 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill8 = new A.SolidFill();

    A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 15000 };
    A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 85000 };

    schemeColor8.Append(luminanceModulation6);
    schemeColor8.Append(luminanceOffset6);

    solidFill8.Append(schemeColor8);
    A.Round round3 = new A.Round();

    outline8.Append(solidFill8);
    outline8.Append(round3);
    A.EffectList effectList8 = new A.EffectList();

    chartShapeProperties8.Append(noFill9);
    chartShapeProperties8.Append(outline8);
    chartShapeProperties8.Append(effectList8);

    TextProperties textProperties4 = new TextProperties();
    A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ListStyle listStyle4 = new A.ListStyle();

    A.Paragraph paragraph4 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill9 = new A.SolidFill();

    A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 65000 };
    A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 35000 };

    schemeColor9.Append(luminanceModulation7);
    schemeColor9.Append(luminanceOffset7);

    solidFill9.Append(schemeColor9);
    A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties4.Append(solidFill9);
    defaultRunProperties4.Append(latinFont4);
    defaultRunProperties4.Append(eastAsianFont4);
    defaultRunProperties4.Append(complexScriptFont4);

    paragraphProperties4.Append(defaultRunProperties4);
    A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph4.Append(paragraphProperties4);
    paragraph4.Append(endParagraphRunProperties4);

    textProperties4.Append(bodyProperties4);
    textProperties4.Append(listStyle4);
    textProperties4.Append(paragraph4);
    CrossingAxis crossingAxis1 = new CrossingAxis() { Val = (UInt32Value)1956507647U };
    Crosses crosses1 = new Crosses() { Val = CrossesValues.AutoZero };
    AutoLabeled autoLabeled1 = new AutoLabeled() { Val = true };
    LabelAlignment labelAlignment1 = new LabelAlignment() { Val = LabelAlignmentValues.Center };
    LabelOffset labelOffset1 = new LabelOffset() { Val = (UInt16Value)100U };
    NoMultiLevelLabels noMultiLevelLabels1 = new NoMultiLevelLabels() { Val = false };

    categoryAxis1.Append(axisId3);
    categoryAxis1.Append(scaling1);
    categoryAxis1.Append(delete1);
    categoryAxis1.Append(axisPosition1);
    categoryAxis1.Append(numberingFormat1);
    categoryAxis1.Append(majorTickMark1);
    categoryAxis1.Append(minorTickMark1);
    categoryAxis1.Append(tickLabelPosition1);
    categoryAxis1.Append(chartShapeProperties8);
    categoryAxis1.Append(textProperties4);
    categoryAxis1.Append(crossingAxis1);
    categoryAxis1.Append(crosses1);
    categoryAxis1.Append(autoLabeled1);
    categoryAxis1.Append(labelAlignment1);
    categoryAxis1.Append(labelOffset1);
    categoryAxis1.Append(noMultiLevelLabels1);

    ValueAxis valueAxis1 = new ValueAxis();
    AxisId axisId4 = new AxisId() { Val = (UInt32Value)1956507647U };

    Scaling scaling2 = new Scaling();
    Orientation orientation2 = new Orientation() { Val = OrientationValues.MinMax };

    scaling2.Append(orientation2);
    Delete delete2 = new Delete() { Val = false };
    AxisPosition axisPosition2 = new AxisPosition() { Val = AxisPositionValues.Left };

    MajorGridlines majorGridlines1 = new MajorGridlines();

    ChartShapeProperties chartShapeProperties9 = new ChartShapeProperties();

    A.Outline outline9 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill10 = new A.SolidFill();

    A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 15000 };
    A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 85000 };

    schemeColor10.Append(luminanceModulation8);
    schemeColor10.Append(luminanceOffset8);

    solidFill10.Append(schemeColor10);
    A.Round round4 = new A.Round();

    outline9.Append(solidFill10);
    outline9.Append(round4);
    A.EffectList effectList9 = new A.EffectList();

    chartShapeProperties9.Append(outline9);
    chartShapeProperties9.Append(effectList9);

    majorGridlines1.Append(chartShapeProperties9);
    NumberingFormat numberingFormat2 = new NumberingFormat() { FormatCode = "General", SourceLinked = true };
    MajorTickMark majorTickMark2 = new MajorTickMark() { Val = TickMarkValues.None };
    MinorTickMark minorTickMark2 = new MinorTickMark() { Val = TickMarkValues.None };
    TickLabelPosition tickLabelPosition2 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };

    ChartShapeProperties chartShapeProperties10 = new ChartShapeProperties();
    A.NoFill noFill10 = new A.NoFill();

    A.Outline outline10 = new A.Outline();
    A.NoFill noFill11 = new A.NoFill();

    outline10.Append(noFill11);
    A.EffectList effectList10 = new A.EffectList();

    chartShapeProperties10.Append(noFill10);
    chartShapeProperties10.Append(outline10);
    chartShapeProperties10.Append(effectList10);

    TextProperties textProperties5 = new TextProperties();
    A.BodyProperties bodyProperties5 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ListStyle listStyle5 = new A.ListStyle();

    A.Paragraph paragraph5 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill11 = new A.SolidFill();

    A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 65000 };
    A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 35000 };

    schemeColor11.Append(luminanceModulation9);
    schemeColor11.Append(luminanceOffset9);

    solidFill11.Append(schemeColor11);
    A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties5.Append(solidFill11);
    defaultRunProperties5.Append(latinFont5);
    defaultRunProperties5.Append(eastAsianFont5);
    defaultRunProperties5.Append(complexScriptFont5);

    paragraphProperties5.Append(defaultRunProperties5);
    A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph5.Append(paragraphProperties5);
    paragraph5.Append(endParagraphRunProperties5);

    textProperties5.Append(bodyProperties5);
    textProperties5.Append(listStyle5);
    textProperties5.Append(paragraph5);
    CrossingAxis crossingAxis2 = new CrossingAxis() { Val = (UInt32Value)1956592255U };
    Crosses crosses2 = new Crosses() { Val = CrossesValues.AutoZero };
    CrossBetween crossBetween1 = new CrossBetween() { Val = CrossBetweenValues.Between };

    valueAxis1.Append(axisId4);
    valueAxis1.Append(scaling2);
    valueAxis1.Append(delete2);
    valueAxis1.Append(axisPosition2);
    valueAxis1.Append(majorGridlines1);
    valueAxis1.Append(numberingFormat2);
    valueAxis1.Append(majorTickMark2);
    valueAxis1.Append(minorTickMark2);
    valueAxis1.Append(tickLabelPosition2);
    valueAxis1.Append(chartShapeProperties10);
    valueAxis1.Append(textProperties5);
    valueAxis1.Append(crossingAxis2);
    valueAxis1.Append(crosses2);
    valueAxis1.Append(crossBetween1);

    ShapeProperties shapeProperties1 = new ShapeProperties();
    A.NoFill noFill12 = new A.NoFill();

    A.Outline outline11 = new A.Outline();
    A.NoFill noFill13 = new A.NoFill();

    outline11.Append(noFill13);
    A.EffectList effectList11 = new A.EffectList();

    shapeProperties1.Append(noFill12);
    shapeProperties1.Append(outline11);
    shapeProperties1.Append(effectList11);

    plotArea1.Append(layout1);
    plotArea1.Append(barChart1);
    plotArea1.Append(categoryAxis1);
    plotArea1.Append(valueAxis1);
    plotArea1.Append(shapeProperties1);

    Legend legend1 = new Legend();
    LegendPosition legendPosition1 = new LegendPosition() { Val = LegendPositionValues.Bottom };
    Overlay overlay2 = new Overlay() { Val = false };

    ChartShapeProperties chartShapeProperties11 = new ChartShapeProperties();
    A.NoFill noFill14 = new A.NoFill();

    A.Outline outline12 = new A.Outline();
    A.NoFill noFill15 = new A.NoFill();

    outline12.Append(noFill15);
    A.EffectList effectList12 = new A.EffectList();

    chartShapeProperties11.Append(noFill14);
    chartShapeProperties11.Append(outline12);
    chartShapeProperties11.Append(effectList12);

    TextProperties textProperties6 = new TextProperties();
    A.BodyProperties bodyProperties6 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ListStyle listStyle6 = new A.ListStyle();

    A.Paragraph paragraph6 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill12 = new A.SolidFill();

    A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 65000 };
    A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 35000 };

    schemeColor12.Append(luminanceModulation10);
    schemeColor12.Append(luminanceOffset10);

    solidFill12.Append(schemeColor12);
    A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties6.Append(solidFill12);
    defaultRunProperties6.Append(latinFont6);
    defaultRunProperties6.Append(eastAsianFont6);
    defaultRunProperties6.Append(complexScriptFont6);

    paragraphProperties6.Append(defaultRunProperties6);
    A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph6.Append(paragraphProperties6);
    paragraph6.Append(endParagraphRunProperties6);

    textProperties6.Append(bodyProperties6);
    textProperties6.Append(listStyle6);
    textProperties6.Append(paragraph6);

    legend1.Append(legendPosition1);
    legend1.Append(overlay2);
    legend1.Append(chartShapeProperties11);
    legend1.Append(textProperties6);
    PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
    DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };

    ExtensionList extensionList1 = new ExtensionList();

    Extension extension1 = new Extension() { Uri = "{56B9EC1D-385E-4148-901F-78D8002777C0}" };
    extension1.AddNamespaceDeclaration("c16r3", "http://schemas.microsoft.com/office/drawing/2017/03/chart");

    OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16r3:dataDisplayOptions16 xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\"><c16r3:dispNaAsBlank val=\"1\" /></c16r3:dataDisplayOptions16>");

    extension1.Append(openXmlUnknownElement3);

    extensionList1.Append(extension1);
    ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };

    chart1.Append(title1);
    chart1.Append(autoTitleDeleted1);
    chart1.Append(plotArea1);
    chart1.Append(legend1);
    chart1.Append(plotVisibleOnly1);
    chart1.Append(displayBlanksAs1);
    chart1.Append(extensionList1);
    chart1.Append(showDataLabelsOverMaximum1);

    ShapeProperties shapeProperties2 = new ShapeProperties();

    A.SolidFill solidFill13 = new A.SolidFill();
    A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

    solidFill13.Append(schemeColor13);

    A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill14 = new A.SolidFill();

    A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 15000 };
    A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 85000 };

    schemeColor14.Append(luminanceModulation11);
    schemeColor14.Append(luminanceOffset11);

    solidFill14.Append(schemeColor14);
    A.Round round5 = new A.Round();

    outline13.Append(solidFill14);
    outline13.Append(round5);
    A.EffectList effectList13 = new A.EffectList();

    shapeProperties2.Append(solidFill13);
    shapeProperties2.Append(outline13);
    shapeProperties2.Append(effectList13);

    TextProperties textProperties7 = new TextProperties();
    A.BodyProperties bodyProperties7 = new A.BodyProperties();
    A.ListStyle listStyle7 = new A.ListStyle();

    A.Paragraph paragraph7 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties();
    A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties();

    paragraphProperties7.Append(defaultRunProperties7);
    A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph7.Append(paragraphProperties7);
    paragraph7.Append(endParagraphRunProperties7);

    textProperties7.Append(bodyProperties7);
    textProperties7.Append(listStyle7);
    textProperties7.Append(paragraph7);

    ExternalData externalData1 = new ExternalData() { Id = "rId3" };
    AutoUpdate autoUpdate1 = new AutoUpdate() { Val = false };

    externalData1.Append(autoUpdate1);

    chartSpace1.Append(date19041);
    chartSpace1.Append(editingLanguage1);
    chartSpace1.Append(roundedCorners1);
    chartSpace1.Append(alternateContent1);
    chartSpace1.Append(chart1);
    chartSpace1.Append(shapeProperties2);
    chartSpace1.Append(textProperties7);
    chartSpace1.Append(externalData1);
    return chartSpace1;
}
