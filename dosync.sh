#!/bin/ksh
# Synchronises three versions of WordSupport.dotm, by copying the latest version over the other two.

#THREE=a b c

NEWEST=$(ls -t "/Users/charles/Documents/Custom Office Templates/WordSupport.dotm" "/Users/charles/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/WordSupport.dotm" "/Users/charles/Dropbox/Dev/WordSupport/WordSupport.dotm"  | head -1)

echo Synchronising from: $NEWEST
cp "$NEWEST" "/Users/charles/Documents/Custom Office Templates/WordSupport.dotm"
cp "$NEWEST" "/Users/charles/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/WordSupport.dotm"
cp "$NEWEST" "/Users/charles/Dropbox/Dev/WordSupport/WordSupport.dotm"



