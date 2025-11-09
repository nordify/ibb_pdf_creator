#!/bin/zsh
CERT_ID="Developer ID Application: Nordify UG (haftungsbeschrankt) (JG7CFYYV5B)"
sudo rm -rf *.dmg
sudo rm -f *.pdf
sudo rm -f *.zip
sudo rm -rf build/
sudo rm -rf dist/
sudo rm -rf .eggs/
rm -f .DS_Store

pip install -r requirements.txt
pip install pyinstaller
pyinstaller --windowed --name "PDF Creator" --icon=resources/icon.icns --add-data "resources:resources" pdf_creator.py

xattr -cr "dist/PDF Creator.app"

# Make sure entitlements are correct
echo "Checking entitlements file..."
cat entitlements.plist

# Sign the app with proper entitlements
codesign --force --deep --verbose --timestamp --options runtime --entitlements entitlements.plist --sign "$CERT_ID" "dist/PDF Creator.app"

# Sign all dylibs
find "dist/PDF Creator.app/Contents/Frameworks" -name "*.dylib" -exec codesign --force --timestamp --options runtime --entitlements entitlements.plist --verbose -s "$CERT_ID" {} \;

# Verify app signing
echo "Verifying app signing..."
codesign --verify --verbose "dist/PDF Creator.app"

# Create DMG
create-dmg \
  --volname "PDF Creator" \
  --window-size 600 450 \
  --background "resources/background.png" \
  --icon-size 100 \
  --icon "PDF Creator.app" 120 200 \
  --app-drop-link 480 200 \
  "PDF Creator.dmg" \
  "dist/PDF Creator.app"

echo "read 'icns' (-16455) \"resources/icon.icns\";" > icon.rsrc
Rez -append icon.rsrc -o "PDF Creator.dmg"
SetFile -a C "PDF Creator.dmg"
rm icon.rsrc

# Sign the DMG
codesign --force --deep --verbose --timestamp --options runtime --entitlements entitlements.plist --sign "$CERT_ID" "PDF Creator.dmg"

# Create a zip archive for notarization (sometimes more reliable than DMG)
echo "Creating zip archive for notarization..."
ditto -c -k --keepParent "dist/PDF Creator.app" "PDF_Creator.zip"

# Add notarization steps with detailed output
echo "Notarizing app..."
xcrun notarytool submit "PDF_Creator.zip" \
    --keychain-profile "PDF_Creator_Notarization" \
    --wait \
    --output-format json > notarization_info.json

# Display notarization results
echo "Notarization results:"
cat notarization_info.json

# Extract submission ID and status
SUBMISSION_ID=$(cat notarization_info.json | python3 -c "import sys, json; print(json.load(sys.stdin).get('id', ''))")
echo "Submission ID: $SUBMISSION_ID"

# Get detailed log for debugging - only if we have a valid submission ID
if [[ -n "$SUBMISSION_ID" ]]; then
    echo "Getting detailed notarization log..."
    xcrun notarytool log "$SUBMISSION_ID" --keychain-profile "PDF_Creator_Notarization" notarization_log.json
    
    if [[ -f notarization_log.json ]]; then
        echo "Notarization log details:"
        cat notarization_log.json
    else
        echo "Failed to retrieve notarization log"
    fi
fi

# Check if notarization was successful
NOTARIZATION_STATUS=$(cat notarization_info.json | python3 -c "import sys, json; print(json.load(sys.stdin).get('status', ''))")
echo "Notarization status: $NOTARIZATION_STATUS"

if [[ "$NOTARIZATION_STATUS" == "Accepted" ]]; then
    echo "Notarization successful! Stapling ticket..."
    
    # Staple to the app first
    xcrun stapler staple "dist/PDF Creator.app"
    
    # For the DMG, we need to recreate it after notarization
    echo "Recreating DMG after notarization..."
    # Remove the old DMG
    rm -f "PDF Creator.dmg"
    
    # Create a new DMG from the notarized app
    create-dmg \
      --volname "PDF Creator" \
      --window-size 600 450 \
      --background "resources/background.png" \
      --icon-size 100 \
      --icon "PDF Creator.app" 120 200 \
      --app-drop-link 480 200 \
      "PDF Creator.dmg" \
      "dist/PDF Creator.app"
    
    # Add icon to the DMG
    echo "read 'icns' (-16455) \"resources/icon.icns\";" > icon.rsrc
    Rez -append icon.rsrc -o "PDF Creator.dmg"
    SetFile -a C "PDF Creator.dmg"
    rm icon.rsrc
    
    # No need to sign the DMG again as it contains the already notarized app
    
    # Verify stapling
    echo "Verifying stapling..."
    xcrun stapler validate "dist/PDF Creator.app"
    
    echo "Build completed successfully!"
else
    echo "Notarization failed with status: $NOTARIZATION_STATUS"
    if [[ -f notarization_log.json ]]; then
        echo "Check notarization_log.json for details"
    fi
    exit 1
fi