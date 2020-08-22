#!/bin/ksh
# Synchronises three versions of ImageAndTableSupport.dotm, by copying the latest version over the other two.

#THREE=a b c

NEWEST=$(ls -t "/Users/charles/Documents/Custom Office Templates/ImageAndTableSupport.dotm" "/Users/charles/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/ImageAndTableSupport.dotm" "/Users/charles/Dropbox/Dev/WordSupport/ImageAndTableSupport.dotm"  | head -1)

echo Synchronising from: $NEWEST
cp "$NEWEST" "/Users/charles/Documents/Custom Office Templates/ImageAndTableSupport.dotm"
cp "$NEWEST" "/Users/charles/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/ImageAndTableSupport.dotm"
cp "$NEWEST" "/Users/charles/Dropbox/Dev/WordSupport/ImageAndTableSupport.dotm"



