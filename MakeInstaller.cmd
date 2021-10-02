@rem Windows Command Script to create ImageAndTableSupport.MSI

@Rem https://serverfault.com/questions/50085/how-do-you-handle-cmd-does-not-support-unc-paths-as-current-directories
pushd "%~dp0\"
candle ImageAndTableSupport.wxs
light -sice:ICE61 ImageAndTableSupport.wixobj
del *.wix*
popd
@pause
