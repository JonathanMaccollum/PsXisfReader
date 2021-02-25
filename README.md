A collection of scripts for maintaining, pre-processing, and stacking images obtained during and after an astrophotography session.

This repository is home to a powershell module named PsXisfReader.

PsXisfReader.psd1
 - A collection of cmdlets for reading data from xisf files per the [XISF 1.0 spec.](https://pixinsight.com/doc/docs/XISF-1.0-spec/XISF-1.0-spec.html)
 - A collection of cmdlets that can be used to launch PixInsight and invoke a number of workflow procedures for pre-processing (calibrating, debayering, weighting, aligning, and even integrating) images.

Required tools:
- PowerShell 7.1 or newer
- A licensed copy of PixInsight installed on the machine maintaining the scripts
- At least one PixInsight "slot" reserved for use by the scripts that you author or would like to run
- A powershell Editor such as Visual Studio Code (Recommended) with the PowerShell 2020.6.0 Preview plugin enabled (or newer)

Getting Started:
- Launch your preferred powershell editor and run the command ``install-module PsXisfReader`` to install the module.
- Create a folder to create a project for maintaining your own scripts.
- Create a powershell script to author your processing workflow.
- At the top of the script add a line to import the PsXisfReader module using the ``import-module PsXisfReader`` command.
- Explore the available cmdlets using the following command: ``get-command module PsXisfReader``

To reserve one or more Slots in PixInsight run the cmdlet ``Start-PixInsight -PixInsightSlot XXX``  one-time. This will launch an instance of PixInsight using the slot number you specified.  It's recommended to configure a dedicated Swap Storage Directory in Global Preferences for each slot you reserve. When you are done, close that instance of PixInsight.