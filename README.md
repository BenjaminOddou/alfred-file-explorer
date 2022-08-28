<img src="public/logo-dark.png#gh-dark-mode-only" alt="logo-dark" height="60"/>
<img src="public/logo-light.png#gh-light-mode-only" alt="logo-light" height="60"/>

[![made with hearth by Benjamin Oddou](https://img.shields.io/badge/made%20with%20%E2%99%A5%20by-benjamin%20oddou-f4bd41.svg?style=flat)](https://github.com/BenjaminOddou)
[![saythanks](https://img.shields.io/badge/say-thanks-7cd9f6.svg?style=flat)](https://saythanks.io/to/BenjaminOddou)

Welcome to the Filenames Renaming repository: **An Alfred Workflowk** ✨

## ✅ Prerequisite

-  MacOS
- ![Alfred logo](public/alfred_logo.svg) Alfred 5. Note that the [Alfred Powerpack](https://www.alfredapp.com/powerpack/) is required to use workflows.
- ![Excel logo](public/excel_logo.svg) Microsoft Excel

## ⬇️ Installation

1. [Download the workflow](https://github.com/BenjaminOddou/alfred-filenames-renaming/releases/latest)
2. Double click the `.alfredworkflow` file to install

<img src="public/workflow.png" alt="Workflow" width="600"/>

## 📖 Documentation

> 🚨 **Important** 
> 
> This workflow doesn't convert files so take care to preserve original filename extensions

> 💡 **Recommended**
>
> Show filename extensions on Mac (See [Documentation](https://support.apple.com/en-ie/guide/mac-help/mchlp2304/mac) from Apple)
> 
> - In the Finder <img src="public/finder_logo.png" height="12" width="12" alt="Finder logo.png"> on your Mac, choose *Finder > Preferences*, then click *Advanced*.
> - Select *Show all filename extensions*.

### 1. Get all filenames of a selected folder in an Excel file

> Trigger: `getfilenames`

<img src="public/trigger_get_filenames-dark.png#gh-dark-mode-only" alt="Get File Names Sreenshot" width="600"/>
<img src="public/trigger_get_filenames-light.png#gh-light-mode-only" alt="Get File Names Sreenshot" width="600"/>

1. Launch the script with the trigger `getfilenames` and press ↩
2. Select the folder with files you want to grab the names of

<img src="public/get_file_names.gif" alt="Gif Get File Names" width="600"/>

### 2. Rename files of a selected folder from an Excel sheet

> Trigger: `renamefiles`

<img src="public/trigger_rename_files-dark.png#gh-dark-mode-only" alt="Rename Files Screenshot" width="600"/>
<img src="public/trigger_rename_files-light.png#gh-light-mode-only" alt="Rename Files Screenshot" width="600"/>

1. In the Excel file created with the first script, input new names in Column B. 
    > 🚨 **Important** : Don't forget filename extensions. 
    > 
    > 💡 **Recommended** : You can use a formula like `="filename"&C1` in `B1` cell to incorporate filename extensions
2. Save your file on your machine with ⌘S or ⌘⇧S
3. Launch the workflow with the trigger `renamefiles`
4. Select the folder with files to rename
5. Select the Excel file with old and new names

<img src="public/rename_files.gif" alt="Gif Rename files" width="600"/>

## ⚖️ License

[MIT License](LICENSE) © Benjamin Oddou