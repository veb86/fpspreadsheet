@echo off
echo -----------------------------
echo  GENERATING CHART DEMO FILES 
echo -----------------------------
echo.
echo AREA SERIES
areachart_write_demo
areachart_write_demo stacked
areachart_write_demo percentage
areachart_write_demo rotated
areachart_write_demo stacked rotated
areachart_write_demo percentage rotated
echo.
echo BAR SERIES
barchart_write_demo vert
barchart_write_demo horiz
barchart_write_demo vert stacked
barchart_write_demo horiz stacked
barchart_write_demo vert percentage
barchart_write_demo horiz percentage
echo.
echo BUBBLE SERIES
bubblechart_write_demo
echo.
echo PIE SERIES
piechart_write_demo
piechart_write_demo ring
echo.
echo RADAR SERIES
radarchart_write_demo
echo.
echo SCATTER SERIES
scatter_write_demo lin
scatter_write_demo log
scatter_write_demo log-log
scatter_write_demo lin inverted
scatter_write_demo log inverted
scatter_write_demo log-log inverted
echo.
echo SCATTER SERIES WITH ERROR BARS
errorbars_write_demo
errorbars_write_demo range
echo.
echo SCATTER SERIES AND REGRESSION DEMO
regressionchart_write_demo
regressionchart_write_demo rotated
echo.
echo STOCK SERIES
stock_write_demo hlc
stock_write_demo candlestick
stock_write_demo hlc rotated
stock_write_demo candlestick rotated
echo.
echo STOCK SERIES WITH VOLUME
stock_volume_write_demo hlc area
stock_volume_write_demo hlc bar
stock_volume_write_demo hlc line
stock_volume_write_demo candlestick area
stock_volume_write_demo candlestick bar
stock_volume_write_demo candlestick line
stock_volume_write_demo hlc area rotated
stock_volume_write_demo hlc bar rotated
stock_volume_write_demo hlc line rotated
stock_volume_write_demo candlestick area rotated
stock_volume_write_demo candlestick bar rotated
stock_volume_write_demo candlestick line rotated
echo.
