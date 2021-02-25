A collection of scripts for maintaining, pre-processing, and stacking images obtained during and after an astrophotography session.

This repository is home to a powershell module named PsXisfReader and PixInsightPreProcessing.

PsXisfReader.psd1
 - A collection of cmdlets for reading data from xisf files per the [XISF 1.0 spec.](https://pixinsight.com/doc/docs/XISF-1.0-spec/XISF-1.0-spec.html)
 - A collection of cmdlets that can be used to launch PixInsight and invoke a number of workflow procedures for pre-processing (calibrating, debayering, weighting, aligning, and even integrating) images.

Required tools:
- PowerShell 5.0 or newer (7.1 recommended)
- A licensed copy of PixInsight installed on the machine maintaining the scripts
- At least one PixInsight "slot" reserved for use by the scripts that you author or would like to run
- A powershell Editor such as PowerShell ISE or Visual Studio Code (Recommended) with the PowerShell 2020.6.0 Preview plugin enabled (or newer)

Getting Started:
- Clone (optionally fork) this repository to the machine you wish to run scripts with.
- Create a folder to create a project for maintaining your scripts
- Create a driver script for your the workflow you'd like to create with the following structure:
- Import the PsXisfReader.psd1 module (for reading information about xisf files)

To reserve one or more Slots in PixInsight run the cmdlet ``Start-PixInsight -PixInsightSlot XXX``  one-time. This will launch an instance of PixInsight using the slot number you specified.  It's recommended to configure a dedicated Swap Storage Directory in Global Preferences for each slot you reserve. When you are done, close that instance of PixInsight