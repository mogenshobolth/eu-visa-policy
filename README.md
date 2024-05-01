# Purpose

The EU visa policy repository contains source code for generating a dataset on visitor visas applied for, issued and not issued by European Union member states. 

# Overview and where to start

The code for generating the dataset is in Python and consists of three files: 
- Program.py, the main code for constructing the dataset
- geography.py, helper methods for working with cities and countries
- create-output-datasets.py, code for generating simple statistics for e.g. visualisation.

The raw data used to generate the dataset is in the "Input" folder.

The core dataset - visitor-visa-statistics.csv - is in the "Output" folder. The other datasets in the folder are aggregations on top of it. 

The repository contains configuration code for working with the code and data in VS Code using a Python docker development container. 

# How to contribute

If you are interested in contributing to the project please do reach out!

# Data sources and acknowledgements

The visa statistics are pulled from official European Union online sources: The Council Registry (2005-2008) and the Migration and Home Affairs Commission website (2009-2022). The raw data is processed (cleaned-up) as evidenced in the source code. The data for 2005-2012 have been imported relying on earlier data clean-up done in connection with the construction of the European Visa Database (see background section below).

The country classification (income group and regions) are sourced from The World Bank: World Bank Country and Lending Groups: [Country classification dataset](ttps://datahelpdesk.worldbank.org/knowledgebase/articles/906519-world-bank-country-and-lending-groups)

# Background 

The EU visa policy repository builds upon my [PhD thesis](http://etheses.lse.ac.uk/551/) handed in at the [London School of Economics](https://www.lse.ac.uk). The thesis put in place a [European Visa Database](http://www.mogenshobolth.dk/evd/) (EVD) covering the period 2005-2012. 