# :dna: PromKappa 1.0 makes Objective Digital Stains

Version <kbd>2.0</kbd> is available here:
https://github.com/Gagniuc/PromKappa-2.0-makes-Objective-Digital-Stains

Version <kbd>3.0</kbd> is available here:
https://github.com/Gagniuc/PromKappa-3.0-Objective-Digital-Stains-in-VB6

# Description
<kbd>PromKappa 1.0</kbd> is a complete application made in <kbd>Visual Basic 6.0 (VB6)</kbd> that generates <kbd>Objective Digital Stains</kbd> (aka DNA patterns). The main output of the application consists of a series of images that are saved in the <kbd>chart</kbd> (or <kbd>chart_comp</kbd>) folder and which can be later analysed using different methods. Note that if you are familiar with python, then <kbd>VB6</kbd> will come natural to you. Also note that once the application is started, the first step will be to load a FASTA file (the <kbd>Homo sapiens (8515).txt</kbd>) using the <kbd>Open promoter file</kbd> button. In the PromKappa case the <a href="https://github.com/Gagniuc/PromKappa-1.0-makes-Objective-Digital-Stains/blob/main/Homo%20sapiens%20(8515).txt"><kbd>Homo sapiens (8515).txt</kbd></a> contains a series of gene promoters, as the analysis of gene promoters was the main aim of this application.

![screenshot](https://github.com/Gagniuc/PromKappa-1.0-makes-Objective-Digital-Stains/blob/main/img/Prom%20Kappa%20(gene%20promoters%20in%20eukaryotes).gif.PNG)

The compiled version of PromKappa (<kbd>PromKappa.exe</kbd>) will ask for a dependency file called <kbd>msvbvm60.dll</kbd> and possibly other dependency files. These files are present in the <kbd>bin</kbd> folder. The following files are a complete set of dependencies that a regular VB6 app may require:

- <kbd>msvbvm60.DLL</kbd>
- <kbd>VBA6.DLL</kbd>
- <kbd>shlwapi.dll</kbd>
- <kbd>MSCOMCTL.OCX</kbd>
- <kbd>COMDLG32.OCX</kbd>

# Implementations - other
The Objective Digital Stains are also implemented in two scripting languages, from which an entire customised application can be made.

In <kbd>Java Script</kbd>:
https://github.com/Gagniuc/Objective-Digital-Stains

In <kbd>PHP</kbd>:
https://github.com/Gagniuc/Objective-Digital-Stains-in-PHP

# Info on ODSs
 Please read more about DNA patterns (aka Objective Digital Stains) here:
 ```
 Eukaryotic genomes may exhibit up to 10 generic classes of gene promoters: 
 ```
 https://bmcgenomics.biomedcentral.com/articles/10.1186/1471-2164-13-512
 
 ```
 Gene promoters show chromosome-specificity and reveal chromosome territories in humans:
 ```
 https://bmcgenomics.biomedcentral.com/articles/10.1186/1471-2164-14-278
 
 ```
 Algorithms in Bioinformatics: Theory and Implementation:
 ```
 https://www.wiley.com/en-ag/Algorithms+in+Bioinformatics%3A+Theory+and+Implementation-p-9781119697961
 
# References
<i>Gagniuc P.A. and Ionescu-Tirgoviste C.: Eukaryotic genomes may exhibit up to 10 generic classes of gene promoters. BMC Genomics 2012, 13:512.</i>

<i>Ionescu-Tîrgovişte C*, Gagniuc PA*, Guja C (2015) Structural Properties of Gene Promoters Highlight More than Two Phenotypes of Diabetes. PLoS ONE 10(9): e0137950.</i>

<i>Gagniuc P.A. and Ionescu-Tîrgovişte C. Gene promoters show chromosome specificity and reveal chromosome territories in humans, BMC Genomics 2013, 14:278.</i>

<i>Paul A. Gagniuc. Algorithms in Bioinformatics: Theory and Implementation. John Wiley & Sons, Hoboken, NJ, USA, 2021, ISBN: 9781119697961.</i>
