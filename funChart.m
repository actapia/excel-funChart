function fc = funChart(f,output,name,xRange,yRange)
    fp = fplot(f,xRange,'LineWidth',2);
    sheetname = strcat(name," Data");
    xlswrite(output,[transpose(fp.XData) transpose(fp.YData)],sheetname);
    Excel = actxserver('Excel.Application');
    wb = Excel.workbooks.Open(pwd + "/" + output);
    charts = wb.Charts;
    chart = invoke(charts,'Add');
    chart.Name = strcat(name," Chart");
    chart.ChartType = vbaDefs.xlXYScatterSmoothNoMarkers;
    yAxis = invoke(chart,'Axes',int32(XlAxisType.xlValue));
    yAxis.MajorGridlines.Format.Line.Visible = int32(MsoTriState.msoFalse);
    xAxis = invoke(chart,'Axes',int32(XlAxisType.xlCategory));
    xAxis.MinimumScale = xRange(1);
    xAxis.MaximumScale = xRange(2);
    if exist('yRange','var')
        yAxis.MinimumScale = yRange(1);
        yAxis.MaximumScale = yRange(2);
    end
    %chart.Axes.Item(XlAxisType.xlValue)
    invoke(chart,'SetSourceData',Excel.Range(getRange(1,1,length(fp.XData),2,sheetname)));
    invoke(chart.Legend,'Delete');
    invoke(chart,'SetElement',int32(MsoChartElementType.msoElementChartTitleAboveChart));
    chart.ChartTitle.Text = name;
    chart.ChartArea.Format.Fill.Visible = int32(MsoTriState.msoTrue);
    chart.ChartArea.Format.Fill.Transparency = 0;
    chart.ChartArea.Format.Line.Visible = int32(MsoTriState.msoTrue);
    chart.ChartArea.Format.Line.Transparency = 0;
    chart.ChartArea.Format.Line.ForeColor.RGB = 0;
    invoke(chart.ChartArea.Format.Fill,'Solid');
    wb.Save
    Excel.Quit;
    delete(Excel);
end