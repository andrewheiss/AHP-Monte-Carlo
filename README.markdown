# AHP + Monte Carlo 

This Excel workbook (or add-in) builds a simplified AHP model with Monte Carlo simulation. Use it and be amazed. :)

The add-in will work with Excel 2007 and 2010.

Enjoy!


# Installation

The package includes two Excel files:

1. AHP Monte Carlo.**xlsm**: A standalone workbook with the custom ribbon tab and buttons. You can use this to create a new model without having to install the macro permanently.
2. AHP Monte Carlo.**xlam**: An add-in for Excel that adds the custom ribbon tab to Excel permanently--it will be available whenever you open Excel. 

## Install the add-in (not recommended)

If you want to have the AHP Monte Carlo program available whenever you open Excel, do the following:

1. Place the .xlam file somewhere on your computer.
2. In Excel, click on the Office button, choose *Excel Options* (or *Options* in Excel 2010)
3. Select *Add-ins* from the left sidebar and click on "Go" next to the the *Manage: Excel Add-ins* dialog
4. Click on "Browse..." and navigate to the .xlam file
5. You're done!


## Uninstall the add-in

If Excel starts crashing or running slow after you've installed the add-in (or if you just don't want it in Excel all the time), do the following:

1. In Excel, click on the Office button, choose *Excel Options* (or *Options* in Excel 2010)
2. Select *Add-ins* from the left sidebar and click on "Go" next to the the *Manage: Excel Add-ins* dialog
3. Uncheck the AHP Monte Carlo add-in in the list
4. You're done!


# Summary

Thomas Saaty's Analytic Hierarchy Process (AHP) is a well known framework for operations and decision analysis. The AHP allows a decision maker to easily decompose a difficult decision into its component parts and assign weights and rankings to each aspect of the problem. By (1) identifying a decision's objectives and alternatives, (2) weighing each of the objectives against each other in a pairwise table, and (3) ranking each of the alternatives in relation to each objective, a decision maker can create a robust and powerful model for decision analysis. An AHP model can be enhanced by using random "fuzzy" values rather than total averages, which gives the final model additional explanatory power, as these fuzzy weights and rankings can be used in a series of Monte Carlo trials.

While an AHP + Monte Carlo model is a fantastic tool for decision analysis, building the model with Excel is relatively time consuming. Pairwise tables that calculate the sumproduct, normalized values, averages, standard deviations, and random inverse normal distributions require dozens of complicated formulas and extensive formatting. More time can often be spent actually building an AHP + Monte Carlo model than analyzing the model's results. 

"AHP + Monte Carlo" is a VBA add-in that automates the majority of the mundane mechanics behind building an AHP + Monte Carlo model in Excel. Users can now build complicated decision support models by simply filling out a series of forms. Instead of spending time repeatedly setting up difficult formulas, users can focus on the actual dynamics of comparing their objectives and alternatives and making better decisions in general.


# Rare object library error

Occasionally, if Excel crashes while using this macro, or if an update to Office goes awry, exposed ActiveX `.exd` files will be left behind and will cause the following error message:

	Object library invalid or contains references to object definitions that could not be found.

According to a [Microsoft KB entry](http://support.microsoft.com/kb/957924/en-us), the solution for this is to search your hard drive for `*.exd` and delete all occurrences. All the necessary `.exd` files will be recreated automatically the next time you open Excel and run the marco.

Alternatively, open Command Prompt and type these two commands:

	CD \Documents and Settings
	DEL /S /A:H /A:-H *.EXD


# License

AHP + Monte Carlo is free and open source software and is provided under the MIT license

Copyright (C) 2011 Andrew Heiss

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.