#!/bin/bash

# Define your passwords and certificate
KEYCHAIN_PASSWORD="your_keychain_password"  # Replace with your actual keychain password
MACOS_CERTIFICATE_PASSWORD="your_certificate_password"  # Replace with your actual certificate password
MACOS_CERTIFICATE="your_base64_encoded_certificate"  # Replace with your actual base64 encoded certificate

# Delete existing keychain if it exists
security delete-keychain build.keychain || echo "No existing keychain to delete."

# Create a new keychain
echo "Creating keychain..."
security create-keychain -p "$KEYCHAIN_PASSWORD" build.keychain || { echo 'Failed to create keychain'; exit 1; }

# Unlock the keychain
echo "Unlocking keychain..."
security unlock-keychain -p "$KEYCHAIN_PASSWORD" build.keychain || { echo 'Failed to unlock keychain'; exit 1; }

# Set keychain settings
echo "Setting keychain settings..."
security set-keychain-setting -t 3600 -l build.keychain

# List keychains
echo "Listing keychains..."
security list-keychains -s build.keychain

# Set keychain access
echo "Setting keychain access..."
security set-keychain-settings -t 3600 -l build.keychain

# Decode and import the certificate
echo "Decoding and importing certificate..."
echo "$MACOS_CERTIFICATE" | base64 --decode > certificate.p12
security import certificate.p12 -k build.keychain -P "$MACOS_CERTIFICATE_PASSWORD" -T /usr/bin/codesign || { echo 'Failed to import certificate'; exit 1; }

# Set key partition list
echo "Setting key partition list..."
security set-key-partition-list -S apple-tool:,apple:,codesign: -s -k "$MACOS_CERTIFICATE_PASSWORD" build.keychain || { echo 'Failed to set key partition list'; exit 1; }

echo "Keychain setup completed successfully."
