# Schengen visa project

The purpose of the Schengen visa project is to generate insights into the implementation of EUs common short-stay visa policy ('Schengen visa') and in turn position these findings in the wider debate on European migration policy and politics as well as EU implementation studies in general. 

## Where to start

Visualisations of the visa issuings statistics:
- Map of the EU/Schengen states visa issuing consulates 
  - [2020 consular map](https://mogenshobolth.github.io/schengen-visa/html/map-visa-issuing-consultates-2020.html)
- Map of of the EU/Schengen states visa refusal rate (not issued rate) and applications received. Applications are on the same scale so the maps can compared over years.
  - [2020 visa statistics map](https://mogenshobolth.github.io/schengen-visa/html/map-visa-statistics-2020.html)
  - [2019 visa statistics map](https://mogenshobolth.github.io/schengen-visa/html/map-visa-statistics-2019.html)
  - [2018 visa statistics map](https://mogenshobolth.github.io/schengen-visa/html/map-visa-statistics-2018.html)

Open the repo in binder to start exploring the notebooks.

[![Binder](https://mybinder.org/badge_logo.svg)](https://mybinder.org/v2/gh/mogenshobolth/schengen-visa/HEAD)

Clone the repo locally and start analysing in Jupyter Notebooks.

## How to contribute

If you are interested in contributing to the project please do reach out!

## Background 

The project builds upon my [PhD thesis](http://etheses.lse.ac.uk/551/) handed in at the [London School of Economics](https://www.lse.ac.uk). The thesis put in place a [European Visa Database](http://www.mogenshobolth.dk/evd/) (EVD) covering the period 2005-2012. Much has happened since then from a methodological perspective with the growth of Data Science as an independent discipline and the resulting diffusion of software enginnering tools and perspectives to other fields. This project hence also seeks to reconstruct this data using contemporary technologies. 

## Folder structure

**data-science**

The DS (Data Science) folder includes script for visual, descriptive and inferential analysis of the data. 

**data-engineering**

The DE (Data Engineering) folder includes script for data processing and preparation. 

**data/bronze**

Contains raw data without any changes applied

**data/silver**

Contains cleaned-up data with a number of transformations etc. applied to make the data nice and easy to work with.