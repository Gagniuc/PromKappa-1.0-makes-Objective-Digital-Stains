# PromKappa 1.0 makes Objective Digital Stains

Version v2.0 is available here:
https://github.com/Gagniuc/PromKappa-2.0-makes-Objective-Digital-Stains


# Description
PromKappa 1.0 is a complete application made in Visual Basic 6.0 (VB6) that generates Objective Digital Stains (aka DNA patterns). The main output of the application consists of a series of images that are saved in the 'chart' (or 'chart_comp') folder and which can be later analysed using different methods. Note that if you are familiar with python, then VB6 will come natural to you. Also note that once the application is started, the first step will be to load a FASTA file (the 'Homo sapiens (8515).txt') using the 'Open promoter file' button. In the PromKappa case the <a href="https://github.com/Gagniuc/PromKappa-1.0-makes-Objective-Digital-Stains/blob/main/Homo%20sapiens%20(8515).txt">Homo sapiens (8515).txt</a> contains a series of gene promoters, as the analysis of gene promoters was the main aim of this application.

![screenshot](https://github.com/Gagniuc/PromKappa-1.0-makes-Objective-Digital-Stains/blob/main/img/Prom%20Kappa%20(gene%20promoters%20in%20eukaryotes).gif.PNG)

The compiled version of PromKappa (PromKappa.exe) will ask for a dependency file called "msvbvm60.dll" and possibly other dependency files. These files are present in the 'bin' folder. The following files are a complete set of dependencies that a regular VB6 app may require:

- msvbvm60.DLL
- VBA6.DLL
- shlwapi.dll
- MSCOMCTL.OCX
- COMDLG32.OCX

# Implementations - other
The Objective Digital Stains are also implemented in two scripting languages, from which an entire customised application can be made.

In Java Script:
https://github.com/Gagniuc/Objective-Digital-Stains

In PHP:
https://github.com/Gagniuc/Objective-Digital-Stains-in-PHP

# Info on ODSs
 Please read more about DNA patterns (aka Objective Digital Stains) here:
 
 Eukaryotic genomes may exhibit up to 10 generic classes of gene promoters: 
 https://bmcgenomics.biomedcentral.com/articles/10.1186/1471-2164-13-512
 
 and 
 
 Gene promoters show chromosome-specificity and reveal chromosome territories in humans:
 https://bmcgenomics.biomedcentral.com/articles/10.1186/1471-2164-14-278
 
 and
 
 Algorithms in Bioinformatics: Theory and Implementation:
 https://www.wiley.com/en-ag/Algorithms+in+Bioinformatics%3A+Theory+and+Implementation-p-9781119697961
 
# References
<i>Gagniuc P.A. and Ionescu-Tirgoviste C.: Eukaryotic genomes may exhibit up to 10 generic classes of gene promoters. BMC Genomics 2012, 13:512.</i>

<i>Ionescu-Tîrgovişte C*, Gagniuc PA*, Guja C (2015) Structural Properties of Gene Promoters Highlight More than Two Phenotypes of Diabetes. PLoS ONE 10(9): e0137950.</i>

<i>Gagniuc P.A. and Ionescu-Tîrgovişte C. Gene promoters show chromosome specificity and reveal chromosome territories in humans, BMC Genomics 2013, 14:278.</i>

<i>Paul A. Gagniuc. Algorithms in Bioinformatics: Theory and Implementation. John Wiley & Sons, Hoboken, NJ, USA, 2021, ISBN: 9781119697961.</i>
