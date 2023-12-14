@echo off
echo Generating chart demo files:
echo.
echo Bar series...
barchart_write_demo
barchart_write_demo rotated
barchart_2axes_write_demo
barchart_2axes_write_demo rotated
echo.
echo Bubble series...
bubblechart_write_demo
echo.
echo Pie series...
piechart_write_demo
echo.
echo Radar series...
radarchart_write_demo
echo.
echo Scatter series...
scatter_write_demo lin
scatter_write_demo log
scatter_write_demo log-log
echo.
echo Scatter series and regression demo...
regressionchart_write_demo
echo.
echo StockSeries...
stock_write_demo hlc
stock_write_demo candlestick
echo.
echo StockSeries with volume...
stock_volume_write_demo hlc area
stock_volume_write_demo hlc bar
stock_volume_write_demo hlc line
stock_volume_write_demo candlestick area
stock_volume_write_demo candlestick bar
stock_volume_write_demo candlestick line
echo.
