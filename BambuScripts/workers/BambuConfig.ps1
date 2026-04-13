# --- Shared configuration for BambuScripts -----------------------------------
# Single source of truth for grandparent theme names and printer prefixes.
# Dot-source this file from any worker script that needs these values:
#   . "$PSScriptRoot\BambuConfig.ps1"

$script:GpThemes = @(
    'Fantasy', 'Puppies', 'Original', 'Ocean', 'Farm', 'Foodz',
    'StarsAndStripes', 'Spring', 'Prehistoric',
    'Halloween 2025', 'Christmas 2025',
    'RTC', 'Summer', 'Licensing', 'Jungle',
    'Halloween26', 'SciFi', 'Sports', 'KidsCreations', 'Careers'
)

# Printer model prefixes used in folder/file naming (e.g. "X1C_Fantasy")
$script:PrinterPrefixes = @('X1C', 'P2S', 'H2S')

# Design variant tags — prepended to the character name in filenames (no separator)
# and shown as "Tag - Character" in card titles/image previews.
# KC = Keychain.  Add new tags here; both CardQueueEditor and DataExtract pick them up automatically.
$script:Tags = @('KC', 'Big', 'Huge')
