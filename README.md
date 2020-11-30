# Errant Knights
A plugin to count BGS work and send to a google sheet
designed for use by multiple Cmdr's

Compatabile with Python 3 release

# Installation
Download the [latest release](https://github.com/SanSin82/Errant-Knights) of BGS Tally
 - In EDMC, on the Plugins settings tab press the “Open” button. This reveals the plugins folder where this app looks for plugins.
 - Open the .zip archive that you downloaded and move the folder contained inside into the plugins folder.
 - Paste and copy your provide creds into the service_account.json file, replacing the entirety of whats in there.

You will need to re-start EDMC for it to notice the new plugin.

# Usage
Errant Knights counts all the BGS work you do for any faction, in any system.
It is highly reccomended that EDMC is started before ED is lauched as Data is recorded at startup and then when you dock at a station. Not doing this can result in missing data.
The data is now stored on a google sheet linked by following the directions below.
Once the credentials have been set multipe Cmdr's can download the JSON file and will be able to add their work to the day's totals
The plugin can paused in the Errant Knights tab in settings.
