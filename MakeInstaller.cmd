@rem Windows Command Script to create ImageAndTableSupport.MSI

pushd "%~dp0\"
candle ImageAndTableSupport.wxs
light ImageAndTableSupport.wixobj
popd
@pause
