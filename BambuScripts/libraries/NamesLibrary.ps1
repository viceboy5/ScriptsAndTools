# --- Shared configuration for BambuScripts -----------------------------------
# Single source of truth for grandparent theme names and printer prefixes.
# Dot-source this file from any worker script that needs these values:
#   . (Join-Path $PSScriptRoot "..\libraries\NamesLibrary.ps1")

$script:GpThemes = @(
    'Fantasy', 'Puppies', 'Original', 'Ocean', 'Farm', 'Foodz',
    'StarsAndStripes', 'Spring', 'Prehistoric', 'Halloween25', 'Christmas25', 'Valentines26',
    'RTC', 'Artemis', 'Punch', 'SixSeven', 'GlobalSoccer', 'Summer',
    'Jungle', 'Halloween26', 'MothersDay', 'SciFi', 'Sports', 'KidsCreations',
    'Careers', 'Maverik', 'KPop', 'GirlScouts', 'Bluey', 'Hersheys',
    'Wicked', 'Geppettos'
)

$script:PrinterPrefixes = @('X1C', 'P2S', 'H2S')

$script:Tags = @('KC', 'Big', 'Huge', 'High')

$script:TagLabels = @{
    'KC'   = 'Keychain'
    'Huge'   = 'Huge Wig'
    'High'   = 'High Res'
    'Big'   = 'Big Wig'
}
